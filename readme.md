# 预算项目管理与文档生成系统 - README
[![Python Version](https://img.shields.io/badge/Python-3.7%2B-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-Private-red.svg)](LICENSE)

## 一、项目简介
本项目是一款基于Python Tkinter开发的**预算项目管理与文档生成GUI工具**，主要面向工程预算、项目管理场景，提供预算项目的增删改查、数据持久化存储、Excel导出以及Word文档（申请表、会审单）自动生成功能。工具支持跨平台运行（Windows、macOS、麒麟Linux），并针对不同系统做了环境适配。

### 核心功能
1. **项目数据管理**：支持施工/材料项目的新增、删除、修改、查看，支持工程量的快速编辑。
2. **数据持久化**：项目数据自动保存为JSON文件，重启后可加载历史数据。
3. **Excel导出**：筛选工程量大于0的项目，导出为Excel文件（支持xlsx格式）。
4. **Word文档生成**：根据项目数据自动填充申请表、会审单模板，支持图片插入、清单/计划内容自动填写（保留原模板文字，右侧列填充内容）。
5. **跨平台兼容**：适配Windows、macOS、麒麟Linux（银河麒麟/中标麒麟）系统。

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
将项目代码复制到本地目录（如`D:\budget-project-system`），确保项目目录包含主脚本（如`main.py`）、配置文件等。

#### 步骤4：运行脚本
在项目目录下，执行以下命令：
```bash
cd D:\budget-project-system
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
将项目代码复制到本地目录（如`~/Documents/budget-project-system`）。

#### 步骤5：运行脚本
在终端中执行以下命令：
```bash
cd ~/Documents/budget-project-system
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
将项目代码复制到本地目录（如`/home/user/budget-project-system`）。

#### 步骤5：运行脚本
```bash
cd /home/user/budget-project-system
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

## 五、常见问题解决
### 1. tkinter模块缺失
- **Windows**：重新安装Python并勾选“Add Python to PATH”，或通过`pip install tkinter`安装。
- **macOS**：执行`brew install tcl-tk`，并重新安装Python。
- **麒麟Linux**：执行`sudo apt install python3-tk`。

### 2. 依赖库安装失败
- 原因：外网访问慢或网络限制。
- 解决：使用国内PyPI源（如清华源），在`pip install`命令后添加`-i https://pypi.tuna.tsinghua.edu.cn/simple/`。

### 3. Word文档生成后内容覆盖原模板文字
- 原因：模板中“工作量及材料清单”等关键词的单元格位置异常。
- 解决：确保模板中的关键词单独在一个单元格中，工具会在右侧列填充内容；或检查项目代码中的关键词查找逻辑是否匹配模板文字。

### 4. Excel导出失败
- 原因：缺少`openpyxl`库（pandas导出xlsx依赖）。
- 解决：执行`pip install openpyxl`。

### 5. 打包后的可执行文件无法运行
- **Windows**：确保打包时使用`-w`参数，且依赖库已全部打包；若缺失DLL文件，可安装VC++运行库。
- **macOS**：确保打包时关联了tkinter库，或右键可执行文件选择“打开”（绕过系统安全限制）。
- **麒麟Linux**：确保打包时使用系统对应的Python版本，且安装了`libgtk-3-dev`等GUI依赖（`sudo apt install libgtk-3-dev`）。

## 六、项目结构
```
budget-project-system/
├── main.py                # 主程序入口
├── budget_data.json       # 项目数据持久化文件（自动生成）
├── 申请表模板.docx         # 自定义Word模板（需手动准备）
├── 会审单模板.docx         # 自定义Word模板（需手动准备）
├── icon.ico               # Windows图标文件（可选）
├── icon.icns              # macOS图标文件（可选）
└── README.md              # 项目说明文档
```

## 七、许可证
本项目为**私有项目（Private）**，受版权保护。未经项目所有者的书面授权，任何单位或个人不得擅自复制、修改、分发、商用本项目的代码、文档及衍生产品，禁止将本项目用于任何商业或公开场景。

### 授权说明
- 仅项目授权用户可获取项目代码、进行内部使用和二次开发（需遵循授权协议）。
- 如需获取授权，请联系项目所有者。

## 八、致谢
- [tkinter](https://docs.python.org/3/library/tkinter.html)：Python内置GUI库。
- [tkcalendar](https://pypi.org/project/tkcalendar/)：日期选择控件。
- [pandas](https://pandas.pydata.org/)：数据处理库。
- [python-docx](https://python-docx.readthedocs.io/)：Word文档处理库。
- [pyinstaller](https://www.pyinstaller.org/)：Python打包工具。