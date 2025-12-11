# 预算项目管理与文档生成系统 - README
[![Python Version](https://img.shields.io/badge/Python-3.7%2B-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![GitHub Stars](https://img.shields.io/github/stars/your-username/budget-project-system?style=social)](https://github.com/your-username/budget-project-system)
[![GitHub Forks](https://img.shields.io/github/forks/your-username/budget-project-system?style=social)](https://github.com/your-username/budget-project-system)

## 一、项目简介
本项目是一款基于Python Tkinter开发的**预算项目管理与文档生成GUI工具**，主要面向工程预算、项目管理场景，提供预算项目的增删改查、数据持久化存储、Excel导出以及Word文档（申请表、会审单）自动生成功能。工具支持跨平台运行（Windows、macOS、麒麟Linux），并针对不同系统做了环境适配。

### 核心功能
1. **项目数据管理**：支持施工/材料项目的新增、删除、修改、查看，支持工程量的快速编辑。
2. **数据持久化**：项目数据自动保存为JSON文件，重启后可加载历史数据。
3. **Excel导出**：筛选工程量大于0的项目，导出为Excel文件（支持xlsx格式）。
4. **Word文档生成**：根据项目数据自动填充申请表、会审单模板，支持图片插入、清单/计划内容自动填写（保留原模板文字，右侧列填充内容）。
5. **跨平台兼容**：适配Windows、macOS、麒麟Linux（银河麒麟/中标麒麟）系统。

### 开源说明
本项目已从私有项目转为**开源项目**，遵循MIT开源协议，欢迎社区贡献代码、提交Issue或提出优化建议。

## 二、环境依赖
### 1. 基础环境
- Python 3.7+（推荐3.8~3.10，兼容性最佳）
- 操作系统：Windows 10/11、macOS 10.15+、麒麟Linux（Kylin OS）V4/V10

### 2. 依赖库
| 库名          | 用途                  | 版本要求       |
|---------------|-----------------------|----------------|
| `tkinter`     | GUI界面开发           | Python内置（需单独安装系统依赖） |
| `tkcalendar`  | 日期选择控件          | >=1.6.1        |
| `pandas`      | 数据处理与Excel导出   | >=1.3.0        |
| `openpyxl`    | Excel文件读写（xlsx） | >=3.0.0        |
| `python-docx` | Word文档生成与编辑    | >=0.8.11       |
| `pyinstaller` | 可选，打包为可执行文件 | >=5.0.0        |

## 三、不同系统环境部署/运行流程
### 说明
Python为解释型语言，本项目的“编译”实际包含**环境配置、依赖安装、运行脚本**三个步骤，若需生成可执行文件（如`.exe`/`.app`/`.bin`），可通过`pyinstaller`打包，以下流程包含**运行脚本**和**打包可执行文件**两部分。

---

### 1. Windows系统（Win10/11 64位）
#### 步骤1：安装Python
1. 下载Python 3.7+安装包：[Python官方下载](https://www.python.org/downloads/windows/)（选择**Windows x86-64 executable installer**）。
2. 运行安装包，勾选**Add Python 3.x to PATH**（关键，自动配置环境变量），点击**Install Now**。
3. 验证安装：打开**命令提示符（CMD）**或**PowerShell**，执行以下命令，显示版本号则成功：
   ```bash
   python --version  # 或 python3 --version
   pip --version
   ```

#### 步骤2：安装依赖库
在CMD/PowerShell中执行以下命令：
```bash
# 升级pip（可选，推荐）
python -m pip install --upgrade pip

# 安装核心依赖
pip install tkcalendar pandas openpyxl python-docx

# 可选，安装打包工具
pip install pyinstaller
```

#### 步骤3：获取项目代码
```bash
# 克隆仓库（推荐）
git clone https://github.com/your-username/budget-project-system.git
cd budget-project-system

# 或直接下载源码压缩包，解压后进入目录
```

#### 步骤4：运行脚本
在项目目录下，执行以下命令：
```bash
python main.py
```
此时GUI界面会启动，即可使用功能。

#### 步骤5：可选，打包为EXE可执行文件
在项目目录下执行以下命令（打包后生成`dist`目录，内含可执行文件）：
```bash
pyinstaller -F -w -i icon.ico main.py
# 参数说明：
# -F：打包为单个可执行文件
# -w：无控制台窗口（GUI程序推荐）
# -i：指定图标文件（可选，需提前准备.ico文件）
```

---

### 2. macOS系统（macOS 10.15+ 英特尔/Apple Silicon）
#### 注意事项
macOS自带的Python为2.7版本（已废弃），需手动安装Python 3.x；且`tkinter`依赖系统的Tcl/Tk库，部分macOS版本需单独安装。

#### 步骤1：安装Python
##### 方式1：通过官方安装包（推荐）
1. 下载Python 3.7+安装包：[Python官方下载](https://www.python.org/downloads/macos/)（选择对应架构：Intel/Apple Silicon）。
2. 运行安装包，按照提示完成安装（默认安装路径：`/Library/Frameworks/Python.framework/Versions/3.x`）。
3. 验证安装：打开**终端（Terminal）**，执行以下命令：
   ```bash
   python3 --version
   pip3 --version
   ```

##### 方式2：通过Homebrew安装（推荐开发者）
1. 若未安装Homebrew，执行以下命令安装：
   ```bash
   /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
   ```
2. 安装Python：
   ```bash
   brew install python3
   ```

#### 步骤2：安装Tcl/Tk（修复tkinter缺失问题）
若运行时提示`tkinter`缺失，执行以下命令：
```bash
# 方式1：通过Homebrew安装
brew install tcl-tk

# 方式2：重新安装Python并关联Tcl/Tk（Apple Silicon）
brew reinstall python3 --with-tcl-tk
```

#### 步骤3：安装依赖库
在终端中执行以下命令：
```bash
# 升级pip
python3 -m pip install --upgrade pip

# 安装核心依赖
pip3 install tkcalendar pandas openpyxl python-docx

# 可选，安装打包工具
pip3 install pyinstaller
```

#### 步骤4：获取项目代码
```bash
# 克隆仓库
git clone https://github.com/your-username/budget-project-system.git
cd budget-project-system
```

#### 步骤5：运行脚本
在终端中执行以下命令：
```bash
python3 main.py
```

#### 步骤6：可选，打包为macOS App应用
在项目目录下执行以下命令：
```bash
pyinstaller -F -w -i icon.icns main.py
# 参数说明：
# -i：macOS使用.icns图标文件（需提前准备）
# 打包后dist目录下会生成可执行文件，可右键选择“显示包内容”查看结构
```

---

### 3. 麒麟Linux系统（银河麒麟V4/V10、中标麒麟 64位）
#### 注意事项
麒麟Linux为国产操作系统，基于Debian/Ubuntu或CentOS内核，以下流程以**银河麒麟V10（Debian系）**为例；若为CentOS系，需将`apt`替换为`yum`。

#### 步骤1：安装Python 3.7+
1. 打开**终端**，执行以下命令更新软件源：
   ```bash
   sudo apt update && sudo apt upgrade -y
   ```
2. 安装Python 3.7+及相关依赖：
   ```bash
   # 安装Python 3.8（麒麟V10默认已预装，若缺失则执行）
   sudo apt install python3.8 python3.8-dev python3-pip -y

   # 配置Python 3.8为默认版本（可选）
   sudo update-alternatives --install /usr/bin/python3 python3 /usr/bin/python3.8 1
   sudo update-alternatives --config python3  # 选择Python 3.8
   ```
3. 验证安装：
   ```bash
   python3 --version
   pip3 --version
   ```

#### 步骤2：安装tkinter系统依赖
麒麟Linux下`tkinter`需单独安装系统包，执行以下命令：
```bash
sudo apt install python3-tk -y  # 对应Python 3.x的tkinter依赖
```

#### 步骤3：安装依赖库
```bash
# 升级pip
python3 -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple/

# 安装核心依赖（使用清华源加速，麒麟系统外网访问较慢时推荐）
pip3 install tkcalendar pandas openpyxl python-docx -i https://pypi.tuna.tsinghua.edu.cn/simple/

# 可选，安装打包工具
pip3 install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple/
```

#### 步骤4：获取项目代码
```bash
# 克隆仓库
git clone https://github.com/your-username/budget-project-system.git
cd budget-project-system
```

#### 步骤5：运行脚本
```bash
python3 main.py
```

#### 步骤6：可选，打包为Linux可执行文件
```bash
pyinstaller -F -w main.py
# 打包后dist目录下生成可执行文件，直接运行即可：
# ./dist/main
```

## 四、使用说明
### 1. 首次运行
- 首次运行时，项目会自动创建`budget_data.json`文件（数据持久化文件），用于存储项目数据。
- 若需生成Word文档，需提前准备**申请表模板.docx**和**会审单模板.docx**，并在工具中选择模板路径。
- **开源提示**：请勿在`budget_data.json`中存储敏感数据（如企业机密、个人信息），若需存储敏感数据，建议自行添加加密逻辑。

### 2. 功能操作
#### （1）项目管理
- **新增项目**：点击“➕ 新增施工项目/新增材料项目”，填写项目名称、单价、单位、工程量等信息，点击确认后添加。
- **删除项目**：选中表格中的项目，点击“🗑️ 删除选中项目”即可删除。
- **修改项目**：选中表格中的项目，点击“✏️ 修改项目信息”，可编辑所有字段。
- **编辑工程量**：双击表格中的“工程量”列，可快速修改工程量。

#### （2）数据导出
- 点击“📤 导出工程量>0项目到Excel”，选择保存路径，即可导出筛选后的项目数据。

#### （3）文档生成
- 填写项目名称、日期、计划实施周期等核心信息。
- 选择申请表和会审单的Word模板路径。
- （可选）上传支撑图片（最多12张，仅申请表）。
- 点击“🚀 生成申请表+会审单”，选择保存路径，即可生成填充后的Word文档。

### 3. 数据存储
- 项目数据保存在`budget_data.json`文件中，可手动备份该文件以防止数据丢失。
- 若需共享数据，建议清理敏感信息后再进行分享。

## 五、常见问题解决
### 1. tkinter模块缺失
- **Windows**：重新安装Python并勾选“Add Python to PATH”，或通过`pip install tkinter`安装。
- **macOS**：执行`brew install tcl-tk`，并重新安装Python。
- **麒麟Linux**：执行`sudo apt install python3-tk`。

### 2. 依赖库安装失败
- 原因：外网访问慢或网络限制。
- 解决：使用国内PyPI源（如清华源、阿里云源），在`pip install`命令后添加`-i https://pypi.tuna.tsinghua.edu.cn/simple/`。

### 3. Word文档生成后内容覆盖原模板文字
- 原因：模板中“工作量及材料清单”等关键词的单元格位置异常。
- 解决：确保模板中的关键词单独在一个单元格中，工具会在右侧列填充内容；或检查项目代码中的关键词查找逻辑是否匹配模板文字。

### 4. Excel导出失败
- 原因：缺少`openpyxl`库（pandas导出xlsx依赖）。
- 解决：执行`pip install openpyxl`（Windows）或`pip3 install openpyxl`（macOS/Linux）。

### 5. 打包后的可执行文件无法运行
- **Windows**：确保打包时使用`-w`参数，且依赖库已全部打包；若缺失DLL文件，可安装VC++运行库。
- **macOS**：确保打包时关联了tkinter库，或右键可执行文件选择“打开”（绕过系统安全限制）。
- **麒麟Linux**：确保打包时使用系统对应的Python版本，且安装了`libgtk-3-dev`等GUI依赖（`sudo apt install libgtk-3-dev`）。

## 六、项目结构
```
budget-project-system/
├── main.py                # 主程序入口
├── budget_data.json       # 项目数据持久化文件（自动生成，建议添加到.gitignore）
├── 申请表模板.docx         # 自定义Word模板（示例模板可在docs目录下获取）
├── 会审单模板.docx         # 自定义Word模板（示例模板可在docs目录下获取）
├── icon.ico               # Windows图标文件（可选）
├── icon.icns              # macOS图标文件（可选）
├── docs/                  # 文档目录（含模板示例、使用教程）
├── .gitignore             # Git忽略文件（排除敏感数据、临时文件）
└── README.md              # 项目说明文档
```

## 七、贡献指南
### 1. 贡献方式
- **提交Issue**：报告Bug、提出功能需求或优化建议。
- **提交PR**：修复Bug、新增功能、优化代码或文档。
- **完善文档**：补充使用教程、注释或示例模板。

### 2. 贡献流程
1. Fork本仓库到你的GitHub账号。
2. 克隆Fork后的仓库到本地：`git clone https://github.com/your-username/budget-project-system.git`。
3. 创建新分支：`git checkout -b feature/your-feature-name`（功能分支）或`bugfix/your-bug-fix`（修复分支）。
4. 提交代码修改：`git commit -m "feat: 新增XX功能"`（遵循[Conventional Commits](https://www.conventionalcommits.org/)规范）。
5. 推送分支到远程：`git push origin feature/your-feature-name`。
6. 打开GitHub仓库，提交Pull Request（PR），描述修改内容。

### 3. 代码规范
- 遵循PEP 8 Python代码规范。
- 新增功能需添加对应的注释和使用示例。
- 修复Bug需附带测试用例（若适用）。

## 八、行为准则
本项目遵循[Contributor Covenant Code of Conduct](https://www.contributor-covenant.org/version/2/1/code_of_conduct/)，参与项目的贡献者需遵守以下准则：
- 尊重他人，友好沟通，禁止使用攻击性语言。
- 专注于项目改进，不讨论无关话题。
- 对新人友好，耐心解答问题。


## 九、许可证
本项目采用**MIT开源许可证**（MIT License），详见LICENSE文件。以下是许可证核心内容：

```
MIT License

Copyright (c) 2025 [JizhouMao]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## 十、致谢
- [tkinter](https://docs.python.org/3/library/tkinter.html)：Python内置GUI库。
- [tkcalendar](https://pypi.org/project/tkcalendar/)：日期选择控件。
- [pandas](https://pandas.pydata.org/)：数据处理库。
- [python-docx](https://python-docx.readthedocs.io/)：Word文档处理库。
- [pyinstaller](https://www.pyinstaller.org/)：Python打包工具。
- 所有为项目贡献代码、提出建议的社区成员。

---
