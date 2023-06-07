## rust-chemdraw-xlsx

```
(base) ➜  rust-chemdraw-xlsx git:(master) ✗ rust-chemdraw-xlsx -h
rust-chemdraw-xlsx 1.0.0

USAGE:
    rust-chemdraw-xlsx [OPTIONS]

FLAGS:
    -h, --help       Prints help information
    -v, --version    显示版本

OPTIONS:
    -i <input>               输入含有ChemDraw Object的XLSX文件路径
    -o, --output <output>    输出一个新的xlsx, 包含smiles和img
```


> `-i`: 必须的输入, xlsx文件路径

> `-o`: 可选, 后面添加输出文件名

> `25` 上可以直接运行`rust-chemdraw-xlsx`