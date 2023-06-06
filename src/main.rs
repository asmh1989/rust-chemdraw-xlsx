#[allow(unused_imports)]
#[allow(dead_code)]
use std::{
    fs::{self, File},
    io::{Error, ErrorKind},
    path::{Path, PathBuf},
};

use crate::args::Opt;
use regex::Regex;
use shell::Shell;
use structopt::StructOpt;
use uuid::Uuid;
use xlsxwriter::{prelude::FormatAlignment, worksheet::ImageOptions, Format, Workbook};

mod args;
mod config;
mod shell;

fn obj2smiles(obj: &PathBuf) -> String {
    if let Ok(ole_file) = ole::OleFile::from_file_blocking(obj) {
        let d = ole_file.root();
        if let Some(class_id) = d.class_id.clone() {
            if &class_id[..] == "41BA6D21-A02E-11CE-8FD9-0020AFD1F20C" {
                if let Ok(content) = ole_file.open_stream(&["CONTENTS"]) {
                    let cdx = format!("/tmp/{}.cdx", Uuid::new_v4().to_string());
                    let _ = fs::write(&cdx, &content).unwrap();
                    let shell = Shell::new("/tmp");
                    let res = shell.run(&format!("obabel -icdx {} -ocan", &cdx));
                    let _ = std::fs::remove_file(&cdx);
                    if res.is_ok() {
                        let s = res.ok().unwrap();
                        // let res = shell.run(&format!("obabel -:\"{}\" -ocan", s.trim()));
                        // if res.is_ok() {
                        //     let s = res.ok().unwrap();
                        return s.trim().to_string();
                        // }
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

fn to_usize(s: &str) -> usize {
    let re = Regex::new(r"\D").unwrap();
    let r = re.replace_all(s, "").to_string();
    match r.parse::<usize>() {
        Ok(n) => n,
        Err(_) => 0,
    }
}

fn to_new_xlsx(data: &Vec<String>, output: &str) {
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
        let command = &format!("obabel -:\"{}\" -O {}", f.trim(), &img);
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
        sheet.set_row(y, 240., Some(&format1)).unwrap();

        sheet.write_string(y, 1, f, Some(&format1)).unwrap();
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

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        if file.name().starts_with("xl/embeddings/") && file.is_file() {
            // 如果是目录，将其名称保存到dir_entries中
            let target_path = target_dir.join(file.name());
            let _ = fs::create_dir_all(target_path.parent().unwrap());
            let mut output_file = File::create(&target_path).unwrap();
            std::io::copy(&mut file, &mut output_file).unwrap();
        }
    }

    // 读取生成的 cdx文件列表
    let mut entries = fs::read_dir(target_dir.join("xl/embeddings/"))?
        .map(|res| res.map(|e| e.path()))
        .collect::<Result<Vec<_>, Error>>()?;

    // 按字母顺序对文件路径进行排序
    entries.sort_by_key(|a| {
        if let Some(file_name) = a.file_name() {
            let s = file_name.to_string_lossy().to_string();
            to_usize(&s)
        } else {
            0
        }
    });

    let mut vv: Vec<String> = vec![];
    // 输出文件列表
    for entry in entries {
        if let Some(_) = entry.file_name().and_then(|name| name.to_str()) {
            let s = obj2smiles(&entry);
            vv.push(s.clone());
            println!("{}", &s);
        } else {
            return Err(Error::new(ErrorKind::InvalidData, "Invalid file name"));
        }
    }

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
