# 新疆移动网络维修需求工具 (Xinjiang Mobile Maintenance Tool)
[![React Version](https://img.shields.io/badge/React-18-blue.svg)](https://react.dev/)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.x-blue.svg)](https://www.typescriptlang.org/)
[![Tailwind CSS](https://img.shields.io/badge/Tailwind%20CSS-3.x-blue.svg)](https://tailwindcss.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![GitHub Stars](https://img.shields.io/github/stars/mjz1029/xinjiang-mobile-maintenance-tool?style=social)](https://github.com/mjz1029/xinjiang-mobile-maintenance-tool)
[![GitHub Forks](https://img.shields.io/github/forks/mjz1029/xinjiang-mobile-maintenance-tool?style=social)](https://github.com/mjz1029/xinjiang-mobile-maintenance-tool)

这是一个基于 React 开发的数字化工具，专为新疆移动网络维护人员设计。旨在简化和自动化维修项目的预算计算、表格填写以及官方单据（需求申请单、会审单）的生成工作。

## 📖 项目简介

该工具解决了传统 Excel 手工填报效率低、格式易错、计算繁琐的问题。用户只需在界面上勾选维修项目、填写数量、上传现场照片，系统即可自动计算总预算，并一键生成标准格式的 PDF 文件。

### 开源说明
该项目已从内部私有工具转为**开源项目**，遵循 MIT 开源许可证，欢迎社区贡献代码、提交 Issue 或提出优化建议。同时，项目保留了针对新疆移动业务场景的定制化功能，也支持开发者根据其他场景进行二次开发和适配。

**核心亮点：**
- **零依赖运行**：打包后仅为一个独立的 `.html` 文件，无需安装任何环境，双击即可在任意电脑（Windows/Mac）的浏览器中运行。
- **所见即所得**：实时计算预算金额，动态调整项目清单。
- **持久化存储**：常用人员和单位信息只需输入一次，自动保存到浏览器，下次使用直接加载。
- **标准化输出**：生成的 PDF 严格遵循《新疆移动网络维修需求申请单》和《会审单》的格式要求。
- **易扩展**：代码基于 TypeScript 开发，结构清晰，支持新增维修项目类型、自定义 PDF 模板等扩展需求。

## ✨ 主要功能

1. **基础信息管理**：
    - 自动生成日期。
    - 设定项目名称、计划实施周期（天数）。
    - **自定义人员信息**：支持在界面上配置申请人、项目经理、实施方等信息，配置后自动保存。

2. **预算智能计算**：
    - 内置 60+ 项标准维修项目单价库。
    - 支持**增、删、改、查**：用户可搜索项目，修改单价，或添加临时项目。
    - 自动汇总计算总金额。

3. **支撑文件管理**：
    - 支持上传现场图片（最大支持 12 张）。
    - **智能排版**：图片以 3 列网格形式清晰排列，自动适应版面，既节省空间又保证图片内容完整可见。

4. **单据生成 (PDF)**：
    - 一键生成包含两张表格的 PDF 文件。
    - **精确分页**：已优化分页逻辑，消除多余空白页，确保申请单和会审单各占一页。
    - **跨平台兼容**：针对 Windows 和 macOS 系统进行了渲染优化。

## 🛠 技术栈

- **核心框架**: [React 18](https://react.dev/) + [TypeScript](https://www.typescriptlang.org/)
- **构建工具**: [Vite](https://vitejs.dev/)
- **样式库**: [Tailwind CSS](https://tailwindcss.com/)
- **PDF 生成**: `html2pdf.js`
- **图标库**: `lucide-react`
- **打包插件**: `vite-plugin-singlefile` (实现单文件输出)

## 🚀 快速开始

### 1. 获取项目代码
```bash
# 克隆仓库（推荐）
git clone https://github.com/mjz1029/xinjiang-mobile-maintenance-tool.git
cd xinjiang-mobile-maintenance-tool

# 或直接下载源码压缩包，解压后进入目录
```

### 2. 开发环境运行

如果您是开发者，需要修改代码或进行二次开发：

```bash
# 安装依赖（建议使用 npm 或 yarn/pnpm）
npm install
# 或
yarn install
# 或
pnpm install

# 启动本地开发服务器 (默认端口 3003)
npm run dev
```

浏览器访问 `http://localhost:3003` 即可看到实时效果。

### 3. 项目打包 (生成单文件)

要生成分发给最终用户的版本：

```bash
# 执行构建命令
npm run build
```

构建完成后，检查项目根目录下的 **`dist`** 文件夹。里面会有一个 **`index.html`** 文件。
**这就是最终成品**，您可以直接将这个文件发送给使用者，他们无需安装任何软件，直接双击打开即可使用。

## 📂 项目结构

```
.
├── src/
│   ├── components/
│   │   ├── InputSection.tsx    # 输入区域（预算表、图片上传、CRUD逻辑、人员配置）
│   │   └── PrintPreview.tsx    # PDF 打印预览模板（核心排版逻辑）
│   ├── App.tsx                 # 主入口，处理状态管理（LocalStorage）和 PDF 下载逻辑
│   ├── constants.ts            # 预置的维修项目和单价数据
│   ├── types.ts                # TypeScript 类型定义
│   └── index.tsx               # React 挂载点
├── public/                     # 静态资源目录（如图片、字体，可选）
├── index.html                  # HTML 模板 (包含 html2pdf CDN)
├── vite.config.ts              # Vite 配置 (包含端口和单文件打包配置)
├── package.json                # 项目依赖定义
├── tsconfig.json               # TypeScript 配置
├── tailwind.config.js          # Tailwind CSS 配置
├── .gitignore                  # Git 忽略文件（排除依赖、打包产物等）
├── LICENSE                     # MIT 开源许可证文件
├── CONTRIBUTING.md             # 贡献指南（可选）
└── README.md                   # 项目说明文档
```

## ⚠️ 常见问题与注意事项

### 1. PDF 生成时表格显示不全？
**已修复**。我们在 Windows 系统下采用了特殊的渲染策略：
- 使用了固定的 A4 宽度 (`210mm`) 容器。
- 渲染过程在屏幕外（Off-screen）进行，不受当前浏览器窗口滚动条宽度的影响。
- 如果仍然遇到问题，请检查系统显示设置中的缩放比例（建议设置为 100%）。

### 2. 如何修改预置的单价？
- **方式一（代码层面）**：修改 `src/constants.ts` 文件中的 `BUDGET_ITEMS` 数组，重新打包后分发。
- **方式二（界面操作）**：直接在页面上“新增项目”或编辑现有项目，修改后的信息会保存在浏览器本地存储中。

### 3. 如何修改常用人员名单？
- 直接在网页顶部的“常用基础信息配置”区域展开并修改即可，系统会自动保存到浏览器 LocalStorage。
- 若需批量配置默认人员信息，可修改 `src/constants.ts` 中的默认配置项，重新打包后生效。

### 4. 二次开发时如何自定义 PDF 模板？
- 修改 `src/components/PrintPreview.tsx` 文件中的 HTML 结构和样式，该文件是 PDF 生成的核心模板，可根据需求调整排版、添加/删除字段。

## 🤝 贡献指南

### 1. 贡献方式
- **提交 Issue**：报告 Bug、提出功能需求、建议优化方向或文档改进。
- **提交 PR**：修复 Bug、新增功能、优化代码/性能、完善文档或翻译。
- **分享经验**：在社区中分享该工具的使用场景和二次开发案例。

### 2. 贡献流程
1. **Fork 仓库**：点击 GitHub 页面右上角的 `Fork` 按钮，将仓库复制到你的账号下。
2. **克隆仓库**：将 Fork 后的仓库克隆到本地（参考「快速开始」中的步骤）。
3. **创建分支**：基于 `main` 分支创建新分支，命名规范如下：
   - 功能开发：`feature/xxx`（如 `feature/add-new-project-type`）
   - Bug 修复：`bugfix/xxx`（如 `bugfix/fix-pdf-table-render`）
   - 文档改进：`docs/xxx`（如 `docs/update-readme`）
4. **提交修改**：遵循 [Conventional Commits](https://www.conventionalcommits.org/) 规范提交代码，示例：
   ```bash
   git commit -m "feat: 新增维修项目分类筛选功能"
   git commit -m "fix: 修复macOS下PDF图片排版错乱问题"
   ```
5. **推送分支**：将修改后的分支推送到你的 GitHub 仓库。
6. **提交 PR**：在 GitHub 页面上打开 Pull Request，描述修改内容和目的，等待审核。

### 3. 代码规范
- 遵循 TypeScript 编码规范，确保类型定义完整。
- 遵循 React 函数式组件开发规范，优先使用 Hooks 而非 Class 组件。
- 新增功能需添加必要的注释，复杂逻辑建议补充说明。
- 修复 Bug 需附带复现步骤和修复思路（在 PR 中说明）。

## 📜 行为准则
本项目遵循 [Contributor Covenant Code of Conduct](https://www.contributor-covenant.org/version/2/1/code_of_conduct/)（贡献者公约行为准则），参与项目的贡献者需遵守以下核心准则：
- 尊重他人，使用友好、包容的语言进行沟通。
- 专注于项目改进，不讨论无关的政治、宗教等话题。
- 对新人友好，耐心解答问题，提供建设性的反馈。

## 🔒 安全漏洞报告
若发现项目中的安全漏洞（如本地存储敏感信息、XSS 漏洞等），请不要直接提交 Issue（避免漏洞公开），可通过 GitHub Issues 中添加安全标签后提交，项目维护者会及时处理。

## 📜 许可证
本项目采用 **MIT 开源许可证**（MIT License），详见 LICENSE 文件。以下是许可证核心内容：

```
MIT License

Copyright (c) 2025 JizhouMao

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

## 🙏 致谢
- [React 团队](https://react.dev/)：提供优秀的前端框架。
- [Tailwind CSS 团队](https://tailwindcss.com/)：提供高效的 CSS 框架。
- [html2pdf.js 团队](https://github.com/eKoopmans/html2pdf.js)：提供便捷的 PDF 生成工具。
- [lucide-react 团队](https://lucide.dev/)：提供简洁的图标库。
- 新疆移动网络维护人员的需求反馈，为项目功能设计提供了核心方向。
- 所有为项目贡献代码、提出建议的社区成员。

---
