# docx2pdf — Linux 下 Word 转 PDF

一条命令把 Word 文档转成 PDF：

```bash
./docx2pdf 报告.docx          # → 同目录生成 报告.pdf
```

内置纯 Python 引擎（`python-docx` + `reportlab`）开箱即用，支持中英文混排、表格、图片、
页眉页脚页码；若机器上装了 LibreOffice，会自动改用它以获得最高保真度。

> 目前**只在 Linux 上实测过**（Ubuntu 24.04 / Python 3.13）。代码里保留了 macOS、Windows
> 的字体路径，但没有验证过，欢迎 PR。

## 目录结构

```
docx2pdf/
├── docx2pdf                ← 启动脚本（直接运行这个）
├── docx2pdf.py             命令行入口：批量、递归、并行、输出规划、诊断
├── d2p/
│   ├── docxmodel.py        解析 .docx：段落/样式/列表/表格/图片/分节/页眉页脚
│   ├── renderer.py         排版成 PDF：分页、表格、图片、页码域、纵横混排
│   ├── fonts.py            字体发现与注册，按字符自动选中/英文字体
│   ├── engines.py          引擎层：内置引擎 + 无头 LibreOffice
│   └── convert.py          库级 API：convert_file / convert_many
├── tests/
│   ├── test_smoke.py       端到端测试（41 项）
│   ├── test_edge_cases.py  边界情况测试（54 项）
│   ├── make_samples.py     生成中英文示例文档
│   └── samples/            现成的示例 .docx（可直接试转）
├── requirements.txt        运行依赖（MIT + BSD）
├── requirements-dev.txt    测试依赖（含仅测试用的 PyMuPDF）
├── LICENSE                 MIT
└── README.md
```

## 快速开始

```bash
git clone <仓库地址> && cd docx2pdf               # 1. 进入工具目录（或解压下载的压缩包后 cd 进去）
python3 -m pip install -r requirements.txt      # 2. 装依赖（python-docx、reportlab）
./docx2pdf --doctor                             # 3. 自检：应显示 [OK ] engine builtin
./docx2pdf ~/文档/你的文档.docx                  # 4. 转换；路径可以是任意位置，不必放进本目录
```

转换结果默认就放在源文件旁边：

```
转换前             转换后
~/文档/            ~/文档/
├── 报告.docx      ├── 报告.docx
└── 合同.docx      ├── 报告.pdf      ← 新增
                   ├── 合同.docx
                   └── 合同.pdf      ← 新增
```

## 常用命令

```bash
./docx2pdf 报告.docx                         # 单个文件 → 同目录 报告.pdf
./docx2pdf -o out/报告.pdf 报告.docx         # 指定输出文件
./docx2pdf -d pdf/ a.docx b.docx             # 多个文件 → pdf/ 目录
./docx2pdf -r 资料/ -d pdf/ -j 4             # 递归目录树并并行，保留子目录结构
./docx2pdf "docs/*.docx"                     # 通配符（引号防止 shell 提前展开）
./docx2pdf -n -r 资料/ -d pdf/               # 只看会转换哪些文件，不真正执行
./docx2pdf --engine libreoffice 旧文档.doc   # 高保真引擎（需先装 LibreOffice）
./docx2pdf --doctor                          # 环境自检（引擎 + 字体）
```

| 选项 | 说明 |
| --- | --- |
| `-o` / `-d` | 输出文件 / 输出目录（目录输入时保留相对路径） |
| `-r` / `--no-recursive` | 递归子目录（目录输入默认开启）/ 只处理第一层 |
| `--engine auto\|builtin\|libreoffice` | 默认 `auto`：有 LibreOffice 就用它，否则用内置引擎 |
| `--no-clobber` | 跳过已存在的 PDF（默认覆盖） |
| `-j N` | 并行任务数，默认 `min(CPU, 4)` |
| `-n` / `-q` / `-v` | 试运行 / 静默 / 详细输出 |

退出码：`0` 全部成功，`1` 有文件失败，`2` 没有可处理的输入。路径含空格或中文时记得加引号。

## 输入格式

| 格式 | 内置引擎 | 需要 LibreOffice |
| --- | --- | --- |
| `.docx` `.docm` `.dotx` `.dotm` | ✅ | |
| `.doc` `.odt` `.rtf` `.wps` `.txt` | | ✅ |

内置引擎已覆盖：多级标题、粗体/斜体/下划线/删除线/上下标/颜色/荧光笔、超链接、
项目符号与编号列表（含 `一、二、三` 中文编号）、表格（边框、底纹、合并单元格、嵌套、
跨页重复表头）、内嵌图片、分页符、分节（纵向+横向混排）、页眉页脚与 `PAGE`/`NUMPAGES` 页码。

**已知差异**：中文加粗不会合成（系统中文字体通常只有单一字重）；CJK 标点禁则未实现；
脚注/批注/公式/图表/EMF-SVG 矢量图不支持；首页与奇偶页不同的页眉页脚按默认页眉页脚处理。
需要与 Word 完全一致时请用 `--engine libreoffice`。

## 安装 LibreOffice（可选）

只有需要**与 Word 完全一致的排版**，或要转 `.doc/.odt/.rtf` 时才需要它 —— 装好后
`--engine auto` 会自动改用它，用 `./docx2pdf --doctor` 确认。

```bash
sudo apt-get install -y libreoffice-writer    # Debian / Ubuntu
sudo dnf install -y libreoffice-writer        # Fedora / RHEL
sudo pacman -S --needed libreoffice-fresh     # Arch
sudo snap install libreoffice                 # Snap
```

没有 root 权限时用这两种（装完都能被自动发现）：

```bash
conda install -y -c conda-forge libreoffice                       # conda
flatpak install --user flathub org.libreoffice.LibreOffice        # flatpak
```

也可以下载官方 tar 包解压到任意目录，然后告诉工具它在哪：

```bash
export D2P_SOFFICE=/解压目录/LibreOffice/program/soffice
```

## 字体与中文

Word 里的字体名会被映射到本机字体（`Arial→DejaVu Sans`、`宋体→AR PL UMing`、
`微软雅黑→Droid Sans Fallback` 等），并按**字符**选择能画出该字的字体，所以中英文混排
不会出现方块；所有用到的字体内嵌进 PDF。用 `./docx2pdf --fonts` 查看实际匹配结果，
用 `--font-dir DIR` 追加字体目录（或 `D2P_FONT_DIRS` 替换搜索路径）。

## 排查

| 现象 | 处理 |
| --- | --- |
| 中文显示为方块/空白 | `./docx2pdf --fonts` 确认 CJK 字族可用，或用 `--font-dir` 指向中文字体目录 |
| `.doc` 报错 | 内置引擎不读 `.doc`：按上面「安装 LibreOffice」装好，或先把文件另存为 `.docx` |
| 排版与 Word 有差异 | 改用 `--engine libreoffice` |
| 想看失败原因 | 加 `-v`，或 `D2P_TRACEBACK=1 ./docx2pdf ...` |
| 提示「无法创建多进程，改为串行」 | 容器/沙箱禁用了 POSIX 信号量，结果不受影响，只是不并行 |

## 测试与开发

```bash
python3 -m pip install -r requirements-dev.txt   # 测试依赖（含仅测试用的 PyMuPDF）
python3 tests/test_smoke.py        # 端到端：转换质量 + 命令行行为
python3 tests/test_edge_cases.py   # 边界：合并单元格、长表格分页、分节、超大图片等
python3 tests/make_samples.py      # 重新生成 tests/samples/ 下的示例文档
```

> 依赖与许可证：转换只用到 MIT 的 `python-docx` 和 BSD 的 `reportlab`；`fonttools`/`Pillow`
> 可选。测试用的 PyMuPDF 是 AGPL-3.0，仅用于读取生成的 PDF，不参与转换。字体全部内嵌自
> 系统字体，本项目不分发任何字体文件。

也可以在 Python 里直接用：

```python
from d2p import convert_file

result = convert_file("报告.docx", "报告.pdf", engine="auto")
print(result.ok, result.engine, result.pages, result.error)
```

## 许可证

[MIT](LICENSE)。转换不依赖任何 copyleft 组件；测试用的 PyMuPDF 是 AGPL-3.0，仅用于校验输出。
