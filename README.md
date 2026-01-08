# AI Copilot - Word 智能写作助手

一个类似 Microsoft Copilot 的 Word 原生插件，使用 Office.js 开发，支持 AI 文本分析、内容生成和 Excel 数据分析。

## 功能特性

### 1. 📝 选中文本分析
- 选中 Word 文档中的任意文字
- AI 自动分析文本质量、结构和逻辑
- 提供改进建议和优化方案

### 2. ✨ AI 生成内容
- 输入需求描述（prompt）
- AI 根据需求生成相应内容
- 一键插入到 Word 文档

### 3. 📄 Excel 数据分析
- 上传 Excel 文件
- AI 分析数据并生成报告
- 自动插入分析结果到文档

## 快速开始

### 安装依赖
```bash
npm install
```

### 开发模式
```bash
npm run dev-server
```

在另一个终端启动插件：
```bash
npm start
```

### 构建生产版本
```bash
npm run build
```

## 项目结构

```
├── assets/                 # 图标资源
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html  # 侧边栏 UI
│   │   ├── taskpane.js    # 核心逻辑
│   │   └── taskpane.css   # 样式文件
│   └── commands/          # 命令处理
├── manifest.xml           # 插件配置文件
├── package.json
└── webpack.config.js
```

## 核心 API 说明

### 1. 读取选中文本
```javascript
async function getSelectedText() {
  return await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();
    return selection.text;
  });
}
```

### 2. 写入内容到文档
```javascript
async function writeContentToWord(text) {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertParagraph(text, Word.InsertLocation.end);
    await context.sync();
  });
}
```

## 集成真实 AI API

当前代码使用模拟数据，需要替换为真实的 AI API。

### 方案 1：使用 OpenAI API

在 `taskpane.js` 中修改 `callAIAnalysis` 函数：

```javascript
async function callAIAnalysis(text) {
  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer YOUR_API_KEY'
    },
    body: JSON.stringify({
      model: 'gpt-4',
      messages: [
        {
          role: 'system',
          content: '你是一个专业的文本分析助手，请分析用户提供的文本并给出改进建议。'
        },
        {
          role: 'user',
          content: text
        }
      ]
    })
  });
  
  const data = await response.json();
  return data.choices[0].message.content;
}
```

### 方案 2：使用 Azure OpenAI

```javascript
async function callAIAnalysis(text) {
  const endpoint = 'YOUR_AZURE_ENDPOINT';
  const apiKey = 'YOUR_AZURE_API_KEY';
  
  const response = await fetch(`${endpoint}/openai/deployments/YOUR_DEPLOYMENT/chat/completions?api-version=2024-02-15-preview`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'api-key': apiKey
    },
    body: JSON.stringify({
      messages: [
        { role: 'system', content: '你是一个专业的文本分析助手' },
        { role: 'user', content: text }
      ]
    })
  });
  
  const data = await response.json();
  return data.choices[0].message.content;
}
```

### 方案 3：使用自建后端

推荐方式：创建一个后端服务来处理 AI 请求，避免在前端暴露 API Key。

```javascript
async function callAIAnalysis(text) {
  const response = await fetch('https://your-backend.com/api/analyze', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ text })
  });
  
  const data = await response.json();
  return data.analysis;
}
```

## Excel 文件解析

要真正解析 Excel 文件，需要安装 `xlsx` 库：

```bash
npm install xlsx
```

然后在 `taskpane.js` 中使用：

```javascript
import * as XLSX from 'xlsx';

function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      
      // 读取第一个工作表
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(firstSheet);
      
      resolve({
        name: file.name,
        sheets: workbook.SheetNames,
        data: jsonData
      });
    };
    
    reader.onerror = () => reject(new Error("文件读取失败"));
    reader.readAsArrayBuffer(file);
  });
}
```

## 安全建议

1. **不要在前端代码中硬编码 API Key**
2. 使用环境变量或后端代理来管理密钥
3. 实施请求频率限制
4. 验证用户输入，防止注入攻击
5. 使用 HTTPS 进行所有 API 通信

## 调试技巧

1. 打开 Word 开发者工具：
   - Windows: `Ctrl + Shift + I`
   - Mac: `Cmd + Option + I`

2. 查看控制台日志：
   ```javascript
   console.log("调试信息", data);
   ```

3. 使用 Office.js 错误处理：
   ```javascript
   try {
     await Word.run(async (context) => {
       // 你的代码
     });
   } catch (error) {
     console.error("错误详情:", error);
   }
   ```

## 常见问题

### Q: 插件无法加载？
A: 检查 manifest.xml 中的 URL 是否正确，确保开发服务器正在运行。

### Q: 无法读取选中文本？
A: 确保在 Word 中选中了文字，并且插件有 ReadWriteDocument 权限。

### Q: AI 响应太慢？
A: 考虑添加加载动画，或使用流式响应（SSE）来逐步显示结果。

## 许可证

MIT License

## 参考资源

- [Office Add-ins 官方文档](https://learn.microsoft.com/office/dev/add-ins/)
- [Word JavaScript API](https://learn.microsoft.com/javascript/api/word)
- [Office.js API 参考](https://learn.microsoft.com/javascript/api/office)
