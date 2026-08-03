# 本地文档脱敏工具（GUI）

为 [ZagooYWX/mask-tool](https://github.com/ZagooYWX/mask-tool) 提供的图形界面封装。

所有文件处理均在本地完成，**不会上传任何内容到云端**。

## 下载即用

### macOS（本机构建版本）

1. 在 Finder 中双击 `本地文档脱敏工具`（单文件可执行程序，约 110 MB）。
2. 若系统提示“无法打开”，请在 Finder 中右键 →“打开”一次即可。
3. 程序已内置 `mask-tool`、jieba 词库与默认配置，**无需安装 Python 或任何依赖**。

### Windows

PyInstaller 不能跨平台编译，Windows `.exe` 需要在 Windows 上另行打包（见下文的“自行打包”）。

## 功能

- 批量选择 `.docx` / `.pdf` / `.xlsx` / `.pptx`。
- 选择 `strict` / `smart` / `aggressive` 三种脱敏模式，`smart` 为默认。
- **用户词库**：界面可直接添加公司名、人名、项目名等自定义敏感词；支持从 TXT 批量导入，并随设置自动保存。
- 输出目录默认与原文件相同，文件名自动追加 `_脱敏` 后缀。
- 实时进度条、逐文件状态、处理完成后弹窗汇总。
- 支持导出映射表，用于后续还原脱敏内容。
- 支持双击结果行直接打开对应输出文件。

## 用户词库（strict 模式的核心）

| 模式 | 自动脱敏规则 |
|------|------|
| `strict` | 主要脱敏**词库匹配**（置信度 0.95）；金额、手机号等高置信度正则项也会被脱敏。 |
| `smart` | 在词库基础上，叠加 jieba NER + 正则，适合日常办公。 |
| `aggressive` | 降低阈值，最大化召回，可能误伤普通词。 |

**建议**：把真实的公司全称、客户姓名、项目名称等维护到"用户词库"中，这样：
- 切换 `strict` 模式时仍有稳定、可控的脱敏效果；
- 所有模式下的脱敏结果都更精准；
- 词库随设置自动保存，重启程序后仍在。

> 词库匹配是**精确包含**（不是模糊匹配），因此建议填入完整名称，例如"北京云端科技有限公司"而不是"云端"。

## 关于 PDF

当前 mask-tool 对 PDF 的处理策略为**仅生成检测报告（JSON）**，不会回写新的 PDF 文件。因此选择 PDF 时，输出的是 `_脱敏_检测报告.json`，GUI 会明确标注“PDF 仅输出检测报告”。

## 脱敏效果说明

- 默认 `smart` 模式启用了轻量 jieba NER，可自动识别并替换**人名、地名、金额、手机号、邮箱、身份证号、日期**等敏感信息。
- 对于**公司名、项目名**，由于 jieba 分词会把长名称拆开，识别效果取决于名称是否被切分为一个完整实体；部分长公司名可能只替换其中的地名片段。
- 所有替换都是可逆的，映射表会随结果一并输出（`xxx_脱敏_映射表.json`），可用于还原。
- 如需更高的召回率，可在 GUI 中选择 `aggressive` 模式；如需更保守，可修改 `assets/mask_tool_config/default.yaml` 的 `thresholds.auto_mask`。

## 自行打包

### macOS

```bash
git clone <本仓库地址>
cd masktool-gui
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 -m PyInstaller --noconfirm masktool_gui.spec
```

`requirements.txt` 已包含 `pyinstaller` 与 `mask-tool`（从 GitHub 安装），
无需再手动安装脱敏引擎。打包完成后会在 `dist/` 下生成单文件可执行程序
`本地文档脱敏工具`。

### Windows

```bat
git clone <本仓库地址>
cd masktool-gui
build.bat
```

打包完成后会在 `dist/` 下生成 `本地文档脱敏工具.exe`。

## 依赖

- 打包/运行本 GUI：标准库 `tkinter` + `PyInstaller`。
- 实际脱敏能力已内嵌在打包后的单文件程序中；开发调试时才需要单独安装 [ZagooYWX/mask-tool](https://github.com/ZagooYWX/mask-tool)：

```bash
git clone https://github.com/ZagooYWX/mask-tool.git
cd mask-tool
pip install -e .
```

## 目录结构

```
masktool-gui/
├── app/
│   ├── __init__.py
│   ├── engine.py      # mask-tool 定位与调用封装
│   ├── settings.py    # 配置持久化
│   └── ui.py          # Tkinter 图形界面
├── main.py            # 程序入口
├── tools/
│   └── make_icon.py   # 生成 assets/app.ico
├── assets/
│   ├── app.ico
│   ├── mask_tool_config/   # 内嵌的 mask-tool 默认词库与阈值配置
│   │   ├── default.yaml
│   │   ├── sample_lexicon.yaml
│   │   └── whitelist.yaml
│   └── version.txt
├── requirements.txt
├── masktool_gui.spec  # PyInstaller 配置
├── build.bat          # Windows 一键打包
└── .github/workflows/build-windows.yml
```

## 协议

MIT
