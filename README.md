# StakeJoy Copilot - AI 驱动的 Word 智能写作助手

<div align="center">

**让论文写作更高效，让学术创作更轻松**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Office.js](https://img.shields.io/badge/Office.js-Latest-blue)](https://learn.microsoft.com/office/dev/add-ins/)
[![JavaScript](https://img.shields.io/badge/JavaScript-ES6+-yellow)](https://www.javascript.com/)
[![Webpack](https://img.shields.io/badge/Webpack-5-blue)](https://webpack.js.org/)

🌐 **在线演示**: [即将上线]

[English](#english) | [中文](#chinese)

</div>

---

## 🎯 项目简介

StakeJoy Copilot 是一个专为学术写作设计的 Word 原生插件，通过 AI 技术帮助用户提升论文质量、降低查重率、优化写作效率。

### 核心价值

- 📝 **智能降重**：AI 深度改写，有效降低查重率，保持原意不变
- ✨ **文本扩写**：一键扩充内容，丰富论述，提升论文字数和深度
- 📄 **格式规范**：自动化论文格式调整，符合各类学术规范
- 🎨 **格式仿写**：学习范文结构，智能生成相似风格内容

### 解决的问题

学术写作中常见的痛点：
- 查重率过高，反复修改耗时费力
- 内容不够充实，难以达到字数要求
- 格式调整繁琐，不符合投稿规范
- 缺乏写作参考，难以把握学术风格

StakeJoy Copilot 通过 AI 技术，让学术写作变得更加高效和专业。

---

## 🚀 快速开始

### 系统要求

- Windows 10/11 或 macOS
- Microsoft Word 2016 或更高版本
- Node.js 14+
- npm / pnpm

### 本地开发

#### 安装步骤

```bash
# 克隆仓库
git clone [仓库地址]
cd StakeJoyCopilot

# 安装依赖
npm install

# 启动开发服务器（终端 1）
npm run dev-server

# 启动 Word 插件（终端 2）
npm start
```

Word 会自动打开并加载插件到侧边栏。

### 快速测试

1. 在 Word 文档中输入一段文字
2. 选中文字，点击"智能降重"
3. 查看 AI 改写结果
4. 点击"应用到文档"完成替换

详细使用指南：[QUICKSTART.md](QUICKSTART.md)

---

## 📖 核心功能

### 1. 智能降重 🔍

**功能描述**：
- 选中需要降重的文本
- AI 深度理解语义，进行智能改写
- 保持原意的同时，大幅降低重复率

**技术亮点**：
- 多层次语义分析
- 同义词智能替换
- 句式结构重组
- 段落逻辑优化

**使用场景**：
- 论文查重率过高需要修改
- 引用内容需要转述
- 多篇文献整合去重

### 2. 文本扩写 ✍️

**功能描述**：
- 选中简短的文本段落
- AI 自动扩充内容，增加论述深度
- 保持学术风格和逻辑连贯性

**扩写策略**：
- 添加论据和案例
- 补充理论支撑
- 扩展分析维度
- 丰富表达方式

**使用场景**：
- 论文字数不足
- 论述不够充分
- 需要增加内容深度

### 3. 论文格式化 📋

**功能描述**：
- 一键调整论文格式
- 支持多种学术规范（APA、MLA、Chicago 等）
- 自动处理标题、引用、参考文献格式

**格式化内容**：
- 标题层级和样式
- 段落间距和缩进
- 图表编号和标注
- 参考文献格式

**使用场景**：
- 投稿前格式调整
- 符合学校/期刊要求
- 统一文档风格

### 4. 范文格式仿写 🎨

**功能描述**：
- 上传参考范文
- AI 学习范文的写作风格和结构
- 根据用户主题生成相似风格的内容

**仿写维度**：
- 文章结构布局
- 段落组织方式
- 语言表达风格
- 论证逻辑模式

**使用场景**：
- 学习优秀论文写作
- 保持团队写作风格统一
- 快速生成初稿框架

---

## 🏗️ 技术架构

### 前端技术栈

- **核心库**: Office.js (Word JavaScript API)
- **开发语言**: JavaScript ES6+
- **UI 框架**: HTML5 + CSS3
- **构建工具**: Webpack 5
- **包管理**: npm

### AI 服务集成

- **支持平台**: OpenAI、Azure OpenAI、自定义后端
- **模型**: GPT-4、GPT-3.5-turbo
- **API 封装**: 统一服务层设计
- **安全机制**: 后端代理，API Key 保护

### 核心模块

1. **文本分析引擎**
   - 语义理解
   - 结构分析
   - 查重检测
   - 质量评估

2. **内容生成引擎**
   - 智能改写
   - 内容扩写
   - 风格仿写
   - 格式转换

3. **格式处理引擎**
   - 样式识别
   - 格式应用
   - 批量处理
   - 规范检查

4. **Word 交互层**
   - 文本读取
   - 内容插入
   - 格式设置
   - 选区管理

---

## 💡 使用示例

### 示例 1：智能降重

```javascript
// 1. 选中需要降重的文本
const selectedText = await Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.load("text");
  await context.sync();
  return selection.text;
});

// 2. 调用 AI 降重
const rewrittenText = await aiService.rewriteText(selectedText, {
  mode: 'plagiarism-reduction',
  preserveMeaning: true,
  academicStyle: true
});

// 3. 替换原文
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.insertText(rewrittenText, Word.InsertLocation.replace);
  await context.sync();
});
```

### 示例 2：文本扩写

```javascript
// 扩写选中的段落
const expandedText = await aiService.expandText(selectedText, {
  targetLength: 'double',  // 扩充至 2 倍长度
  addExamples: true,       // 添加案例
  addTheory: true          // 补充理论
});
```

### 示例 3：格式化论文

```javascript
// 应用 APA 格式
await formatDocument({
  style: 'APA',
  fontSize: 12,
  lineSpacing: 2.0,
  margins: { top: 1, bottom: 1, left: 1, right: 1 }
});
```

### 示例 4：范文仿写

```javascript
// 上传范文并生成相似内容
const sampleDoc = await readSampleDocument(file);
const generatedContent = await aiService.imitateStyle(sampleDoc, {
  topic: '人工智能在教育中的应用',
  length: 3000
});
```

---

## 📁 项目结构

```
StakeJoyCopilot/
├── manifest.xml           # Office 插件配置文件
├── src/
│   ├── taskpane/         # 侧边栏界面
│   │   ├── taskpane.html # UI 界面
│   │   ├── taskpane.js   # 核心逻辑
│   │   └── taskpane.css  # 样式文件
│   ├── services/         # 服务层
│   │   ├── aiService.js  # AI 服务封装
│   │   ├── wordService.js # Word API 封装
│   │   └── formatService.js # 格式处理
│   ├── utils/            # 工具函数
│   │   ├── textAnalyzer.js  # 文本分析
│   │   └── styleParser.js   # 样式解析
│   └── config.js         # 配置文件
├── assets/               # 图标资源
├── backend-example/      # 后端示例代码
└── webpack.config.js     # Webpack 配置
```

---

## 🧪 测试

### 功能测试流程

1. **降重功能测试**
   - 在 Word 中输入一段学术文本
   - 选中文本，点击"智能降重"
   - 验证改写后的文本语义一致性
   - 使用查重工具对比重复率

2. **扩写功能测试**
   - 输入简短段落（50-100 字）
   - 点击"文本扩写"
   - 检查扩写后的内容质量和连贯性

3. **格式化测试**
   - 准备一篇未格式化的论文
   - 选择目标格式（APA/MLA）
   - 验证格式应用效果

4. **仿写功能测试**
   - 上传一篇范文
   - 输入新主题
   - 检查生成内容的风格相似度

详细测试指南：[测试说明.md](测试说明.md)

---

## 📚 文档

- [项目基本信息](project_info.md) - 团队和项目信息
- [详细项目文档](project_documentation.md) - 技术架构和实现
- [快速开始指南](QUICKSTART.md) - 快速上手教程
- [开发指南](开发指南.md) - 开发者文档
- [测试说明](测试说明.md) - 完整测试流程
- [项目说明](项目说明.md) - 详细功能说明

---

## 🤝 贡献指南

我们欢迎所有形式的贡献！

### 如何贡献

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

### 代码规范

- 使用 TypeScript 进行类型检查
- 遵循 ESLint 规则
- 添加必要的注释
- 编写单元测试

---

## 📄 开源协议

本项目采用 [MIT License](LICENSE)

### 许可说明

- ✅ 商业使用
- ✅ 修改
- ✅ 分发
- ✅ 私人使用
- ⚠️ 需保留版权声明

---

## 🗺️ 路线图

### Q1 2026
- ✅ 核心功能开发完成
- ✅ 智能降重功能
- ✅ 文本扩写功能
- ⏳ 完善文档和测试

### Q2 2026
- 🎯 论文格式化功能完善
- 🎯 范文仿写功能优化
- 🎯 支持更多学术格式规范
- 🎯 后端服务部署

### Q3 2026
- 🎯 多语言支持（中英文互译）
- 🎯 文献引用管理
- 🎯 图表自动生成
- 🎯 协作编辑功能

### Q4 2026
- 🎯 移动端支持
- 🎯 云端同步功能
- 🎯 模板库建设
- 🎯 社区分享平台

---

## 📞 联系方式

- **GitHub**: [仓库链接]
- **Issues**: [问题反馈]
- **邮箱**: [团队邮箱]
- **Discord**: [社区链接]

---

## 🙏 致谢

感谢以下开源项目和社区：

- [Office.js](https://learn.microsoft.com/office/dev/add-ins/) - Microsoft Office 插件开发框架
- [OpenAI](https://openai.com/) - 强大的 AI 语言模型
- [Webpack](https://webpack.js.org/) - 模块打包工具
- [Office Add-ins Community](https://github.com/OfficeDev) - Office 插件开发社区
- 上海开源信息技术协会 - 提供比赛平台

---

## ⚠️ 免责声明

本项目仅用于学习和研究目的。使用本项目时请注意：

1. AI 生成的内容仅供参考，请自行审核和修改
2. 请遵守学术诚信原则，合理使用 AI 辅助工具
3. 不同学校和期刊对 AI 使用有不同规定，请遵守相关规范
4. 使用本项目造成的任何后果，开发团队不承担责任
5. 请遵守当地法律法规和学术道德规范

---

## 📈 项目状态

- ✅ 智能降重功能完成
- ✅ 文本扩写功能完成
- ⏳ 论文格式化开发中
- ⏳ 范文仿写开发中
- ⏳ 后端服务部署中
- ⏳ 文档完善中

---

<div align="center">

**让学术写作更高效，让论文创作更轻松！** 🚀

Made with ❤️ by StakeJoy Team

</div>
