
# PPT 生成工具

一个基于 Python 的自动化 PowerPoint 生成工具，支持从数据中提取信息并创建幻灯片。

## 📁 项目结构

```
├── generate_ppt_new.py         # 主程序源码
├── requirements.txt            # 所需依赖库
└── dist/
    └── generate_ppt_new.exe    # 编译好的可执行文件
```

## 📦 安装依赖

如需在本地运行源代码，请先安装依赖：

```
pip install -r requirements.txt
```

## 🚀 使用方法

### 1. 使用源码运行

确保已安装所有依赖，运行主脚本：

```
python generate_ppt_new.py
```

### 2. 使用可执行文件（无需安装 Python）

直接运行 `dist/generate_ppt_new.exe` 即可开始使用。

## 🧾 依赖库

详见 `requirements.txt`，主要包括：

* `pandas`：数据处理
* `python-pptx`：PPT 生成
* `pypinyin`：中文转拼音

## 🛠️ 环境建议

* Python 3.8+
* Windows 推荐使用打包好的 `.exe` 文件，无需安装环境

