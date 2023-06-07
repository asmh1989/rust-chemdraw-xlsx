use std::collections::HashMap;
#[allow(unused_imports)]
#[allow(dead_code)]
use std::{
    fs::{self, File},
    io::{Error, ErrorKind},
    path::{Path, PathBuf},
};

use crate::{
    args::Opt,
    xml::{parse_relation_xml, parse_sheet_xml},
};
use shell::Shell;
use structopt::StructOpt;
use uuid::Uuid;
use xlsxwriter::{prelude::FormatAlignment, worksheet::ImageOptions, Format, Workbook};

mod args;
mod config;
mod shell;
mod xml;

fn obj2smiles(obj: &PathBuf) -> String {
    if let Ok(ole_file) = ole::OleFile::from_file_blocking(obj) {
        let d = ole_file.root();
        if let Some(class_id) = d.class_id.clone() {
            if &class_id[..] == "41BA6D21-A02E-11CE-8FD9-0020AFD1F20C" {
                if let Ok(content) = ole_file.open_stream(&["CONTENTS"]) {
                    let cdx = format!("/tmp/{}.cdx", Uuid::new_v4().to_string());
                    let _ = fs::write(&cdx, &content).unwrap();
                    let shell = Shell::new("/tmp");
                    // let res = shell.run(&format!("obabel -icdx {} -ocan -xk", &cdx));
                    let res = shell.run(&format!("obabel -icdx {} -ocan", &cdx));
                    let _ = std::fs::remove_file(&cdx);
                    if res.is_ok() {
                        let s = res.ok().unwrap();
                        return s.trim().to_string();
                    }

                    log::info!("{} error!!", obj.display())
                }
            }
        }
    } else {
        log::info!("{} is not ole file!!", obj.display());
    }
    "".to_owned()
}

pub fn file_exist(path: &str) -> bool {
    std::fs::metadata(path).is_ok()
}

#[derive(Default, Clone)]
struct Cell {
    pub smiles: String,
    pub file_name: String,
    pub row: usize,
    #[allow(dead_code)]
    pub col: usize,
}

fn get_cell(
    entry: &PathBuf,
    relation_map: &HashMap<String, String>,
    id_map: &HashMap<String, (usize, usize)>,
) -> Cell {
    let smiles: String = obj2smiles(&entry);
    let file_name = entry
        .as_path()
        .file_name()
        .unwrap()
        .to_string_lossy()
        .to_string();

    // log::info!("rel = {:?} id= {}", relation_map, &file_name);
    let id = relation_map.get(&file_name).unwrap();
    let r = id_map.get(id).unwrap();
    let row = r.0;
    let col = r.1;
    Cell {
        smiles,
        row,
        col,
        file_name,
    }
}

fn to_new_xlsx(data: &Vec<Cell>, output: &str) {
    let workbook = Workbook::new(output).unwrap();
    let mut sheet = workbook.add_worksheet(None).unwrap();
    let mut format1 = Format::new();
    format1
        .set_align(FormatAlignment::Center)
        .set_vertical_align(xlsxwriter::prelude::FormatVerticalAlignment::VerticalCenter);

    sheet.write_string(0, 0, "struct", Some(&format1)).unwrap();
    sheet.write_string(0, 1, "smiles", Some(&format1)).unwrap();

    sheet.set_column(0, 1, 300., Some(&format1)).unwrap();
    // sheet.set_column(1, 2, 200., Some(&format1)).unwrap();

    let mut y = 1;
    let tmp_dir = format!("/tmp/{}", Uuid::new_v4().to_string());
    let shell = Shell::new(&tmp_dir);
    let _ = fs::create_dir_all(&tmp_dir);

    data.iter().for_each(|f| {
        let img = format!(
            "{}/{}.png",
            &tmp_dir,
            chrono::Local::now().timestamp_nanos()
        );
        let command = &format!("obabel -:\"{}\" -O {}", &f.smiles, &img);
        // log::info!("img = {}, {}", &img, command);
        let _ = shell.run(command);
        sheet
            .insert_image_opt(
                y,
                0,
                &img,
                &ImageOptions {
                    x_offset: 10,
                    y_offset: 10,
                    x_scale: 1.0,
                    y_scale: 1.0,
                },
            )
            .unwrap();
        sheet.set_row(f.row as u32, 240., Some(&format1)).unwrap();

        sheet
            .write_string(f.row as u32, 1, &f.smiles, Some(&format1))
            .unwrap();
        sheet
            .write_string(f.row as u32, 2, &f.file_name, Some(&format1))
            .unwrap();
        y += 1;
    });

    workbook.close().unwrap();

    let _ = std::fs::remove_dir_all(tmp_dir);
}

fn main() -> Result<(), Error> {
    const VERSION: &'static str = env!("CARGO_PKG_VERSION");

    let opt: Opt = Opt::from_args();

    // 打印版本
    if opt.version {
        println!("{}", VERSION);
        return Ok(());
    }

    crate::config::init_config();

    let path = opt.input.map_or_else(|| "".to_string(), |f| f);
    let output = opt.output.map_or_else(|| "".to_string(), |f| f);

    if !file_exist(&path) {
        log::info!("{} 输入文件不存在!", path);
        return Ok(());
    }

    let zipfile = std::fs::File::open(path).unwrap();

    let mut archive = zip::ZipArchive::new(zipfile).unwrap();
    let tmp_dir = format!("/tmp/{}", Uuid::new_v4().to_string());
    // let tmp_dir = format!("data/tmp");

    let target_dir = Path::new(&tmp_dir);
    let _ = fs::create_dir_all(&tmp_dir);

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        if file.name().starts_with("xl/embeddings/") && file.is_file() {
            // 如果是目录，将其名称保存到dir_entries中
            let target_path = target_dir.join(file.name());
            let _ = fs::create_dir_all(target_path.parent().unwrap());
            let mut output_file = File::create(&target_path).unwrap();
            std::io::copy(&mut file, &mut output_file).unwrap();
        } else if file.name().starts_with("xl/worksheets/sheet1.xml") {
            let target_path = target_dir.join("sheet1.xml");
            let mut output_file = File::create(&target_path).unwrap();
            std::io::copy(&mut file, &mut output_file).unwrap();
        } else if file
            .name()
            .starts_with("xl/worksheets/_rels/sheet1.xml.rels")
        {
            let target_path = target_dir.join("sheet1_relation.xml");
            let mut output_file = File::create(&target_path).unwrap();
            std::io::copy(&mut file, &mut output_file).unwrap();
        }
    }

    // 读取生成的 cdx文件列表
    let entries = fs::read_dir(target_dir.join("xl/embeddings/"))?
        .map(|res| res.map(|e| e.path()))
        .collect::<Result<Vec<_>, Error>>()?;

    let mut vv: Vec<Cell> = vec![];
    let releation_map = parse_relation_xml(target_dir.join("sheet1_relation.xml"));
    let id_map = parse_sheet_xml(target_dir.join("sheet1.xml"));
    // 输出文件列表
    for entry in entries {
        if let Some(_) = entry.file_name().and_then(|name| name.to_str()) {
            let s = get_cell(&entry, &releation_map, &id_map);
            vv.push(s.clone());
            // println!("{}, {}, {}", &s.smiles, s.row, s.col);
        } else {
            return Err(Error::new(ErrorKind::InvalidData, "Invalid file name"));
        }
    }

    vv.sort_by_key(|f| f.row);

    vv.iter().for_each(|s| {
        // println!("{}, {}, {}, {}", &s.smiles, s.row, s.col, &s.file_name);
        println!("{}", &s.smiles);
    });

    let _ = fs::remove_dir_all(tmp_dir);

    if !output.is_empty() {
        if file_exist(&output) {
            let _ = fs::remove_file(&output);
        }
        to_new_xlsx(&vv, &output);

        log::info!("输出文件 = {}", &output);
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_obj2smiles() {
        crate::config::init_config();
        let file = Path::new("data/tmp/xl/embeddings/oleObject9.bin");
        // let file = Path::new("demo/cdxFiles/oleObject8.bin.cdx");

        log::info!("bin to smiles = {}", obj2smiles(&file.to_path_buf()));
    }
}
