use structopt::StructOpt;

#[derive(StructOpt, Debug)]
#[structopt(name = "rust-chemdraw-xlsx")]
pub struct Opt {
    #[structopt(short = "v", long, help = "显示版本")]
    pub version: bool,

    #[structopt(short = "i", help = "输入含有ChemDraw Object的XLSX文件路径")]
    pub input: Option<String>,
    #[structopt(
        long = "output",
        short = "o",
        help = "输出一个新的xlsx, 包含smiles和img"
    )]
    pub output: Option<String>,
    // #[structopt(long = "id", help = "id文件列表")]
    // pub id: Option<String>,

    // #[structopt(long = "csv", help = "输出csv文件格式")]
    // pub csv: bool,
}
