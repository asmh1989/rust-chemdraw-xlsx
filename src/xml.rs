use std::{collections::HashMap, io::BufReader, path::Path};

use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
pub struct Relationships {
    #[serde(rename = "Relationship")]
    pub relationships: Vec<Relationship>,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct Relationship {
    #[serde(rename = "Id")]
    pub id: String,

    #[serde(rename = "Type")]
    pub rel_type: String,

    #[serde(rename = "Target")]
    pub target: String,
}

pub fn parse_relation_xml<P: AsRef<Path>>(p: P) -> HashMap<String, String> {
    let file = std::fs::File::open(p).unwrap();
    let reader = BufReader::new(file);
    let relationships: Relationships = serde_xml_rs::from_reader(reader).unwrap();

    let mut map: HashMap<String, String> = HashMap::new();
    relationships.relationships.iter().for_each(|f| {
        if f.target.ends_with(".bin") {
            map.insert(
                f.target.replace("../embeddings/", "").to_string(),
                f.id.clone(),
            );
        }
    });

    map
}

pub fn parse_sheet_xml<P: AsRef<Path>>(p: P) -> HashMap<String, (usize, usize)> {
    let file = std::fs::File::open(p).unwrap();
    let reader = BufReader::new(file);
    let relationships: Worksheet = serde_xml_rs::from_reader(reader).unwrap();

    // log::info!(
    //     "sheet = {}",
    //     serde_json::to_string_pretty(&serde_json::to_value(&relationships).unwrap()).unwrap()
    // );

    let mut map: HashMap<String, (usize, usize)> = HashMap::new();
    relationships.ole_objects.alt_content.iter().for_each(|f| {
        map.insert(
            f.choice.ole_object.r_id.clone(),
            (
                f.choice.ole_object.object_pr.anchor.from.row,
                f.choice.ole_object.object_pr.anchor.from.col,
            ),
        );
    });

    map
}

#[derive(Debug, Deserialize, Serialize)]
pub struct Worksheet {
    #[serde(rename = "oleObjects")]
    pub ole_objects: OleObjects,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct OleObjects {
    #[serde(rename = "$value")]
    pub alt_content: Vec<AlternateContent>,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct AlternateContent {
    #[serde(rename = "Choice")]
    pub choice: Choice,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct Choice {
    #[serde(rename = "oleObject")]
    pub ole_object: OleObject,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct OleObject {
    #[serde(rename = "progId")]
    pub prog_id: String,

    #[serde(rename = "id")]
    pub r_id: String,

    #[serde(rename = "objectPr")]
    pub object_pr: ObjectPr,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct ObjectPr {
    #[serde(rename = "anchor")]
    pub anchor: Anchor,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct Anchor {
    #[serde(rename = "from")]
    pub from: From,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct From {
    #[serde(rename = "col")]
    pub col: usize,

    #[serde(rename = "row")]
    pub row: usize,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_relation_xml() {
        crate::config::init_config();
        log::info!(
            "result = {:?}",
            &parse_relation_xml("data/tmp/sheet1_relation.xml")
        );
    }

    #[test]
    fn test_parse_sheet_xml() {
        crate::config::init_config();
        log::info!("result = {:?}", &parse_sheet_xml("data/tmp/sheet1.xml"));
        log::info!("result = {:?}", &parse_sheet_xml("demo/sheet1.xml"));
    }

    #[test]
    fn test_path() {
        crate::config::init_config();
        let p = Path::new("data/tmp/sheet1_relation.xml");
        log::info!("file_name = {}", p.file_name().unwrap().to_string_lossy());
    }
}
