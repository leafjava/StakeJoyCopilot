# 快速开始指南

## 第一步：安装依赖

```bash
npm install
```

## 第二步：启动开发服务器

打开第一个终端，运行：

```bash
npm run dev-server
```

等待服务器启动，你会看到类似这样的输出：
```
webpack 5.x.x compiled successfully
Server running at https://localhost:3000/
```

## 第三步：启动 Word 插件

打开第二个终端，运行：

```bash
npm start
```

这会自动打开 Word 并加载你的插件。

## 第四步：测试插件功能

### 测试 1：分析选中文本
1. 在 Word 文档中输入一些文字
2. 选中这些文字
3. 在插件侧边栏点击"分析选中文本"
4. 查看 AI 分析结果

### 测试 2：生成内容
1. 在"AI 生成内容"区域输入需求，例如："写一段关于云计算的介绍"
2. 点击"生成内容"按钮
3. 查看生成结果
4. 点击"插入到文档"将内容添加到 Word

### 测试 3：Excel 分析
1. 准备一个 Excel 文件
2. 在插件中点击"选择文件"
3. 选择你的 Excel 文件
4. 点击"分析 Excel"
5. 查看分析报告

## 第五步：配置真实的 AI API（可选）

当前插件使用模拟数据。要使用真实的 AI 功能：

### 方案 A：使用 OpenAI

1. 复制配置文件：
```bash
copy src\config.example.js src\config.js
```

2. 编辑 `src/config.js`，填入你的 OpenAI API Key：
```javascript
export const AI_CONFIG = {
  provider: 'openai',
  openai: {
    apiKey: 'sk-your-api-key-here',
    model: 'gpt-4',
    // ...其他配置
  }
};
```

3. 重启开发服务器

### 方案 B：使用 Azure OpenAI

1. 复制配置文件（同上）

2. 编辑 `src/config.js`：
```javascript
export const AI_CONFIG = {
  provider: 'azure',
  azure: {
    endpoint: 'https://your-resource.openai.azure.com',
    apiKey: 'your-azure-api-key',
    deploymentName: 'your-deployment-name',
    apiVersion: '2024-02-15-preview'
  }
};
```

3. 重启开发服务器

### 方案 C：使用自建后端（推荐）

1. 创建一个后端服务（Node.js/Python/Java 等）
2. 实现以下 API 端点：
   - POST `/api/analyze` - 文本分析
   - POST `/api/generate` - 内容生成
   - POST `/api/excel-analysis` - Excel 分析

3. 配置 `src/config.js`：
```javascript
export const AI_CONFIG = {
  provider: 'custom',
  custom: {
    baseURL: 'https://your-backend.com/api',
    apiKey: 'your-backend-api-key',
    endpoints: {
      analyze: '/analyze',
      generate: '/generate',
      excelAnalysis: '/excel-analysis'
    }
  }
};
```

## 调试技巧

### 查看控制台日志
在 Word 中按 `F12` 或 `Ctrl + Shift + I` 打开开发者工具。

### 常见错误

**错误：插件无法加载**
- 检查开发服务器是否正在运行
- 确认 manifest.xml 中的 URL 是 `https://localhost:3000`
- 尝试清除 Office 缓存

**错误：无法读取选中文本**
- 确保在 Word 中选中了文字
- 检查插件权限设置

**错误：AI API 调用失败**
- 检查 API Key 是否正确
- 确认网络连接正常
- 查看控制台错误信息

## 构建生产版本

```bash
npm run build
```

构建完成后，文件会输出到 `dist/` 目录。

## 部署到生产环境

1. 将 `dist/` 目录部署到你的 Web 服务器
2. 更新 `manifest.xml` 中的 URL 为生产环境地址
3. 在 Microsoft 365 管理中心上传 manifest.xml

## 需要帮助？

- 查看 [README.md](./README.md) 了解详细文档
- 访问 [Office Add-ins 官方文档](https://learn.microsoft.com/office/dev/add-ins/)
- 查看 [Word JavaScript API 参考](https://learn.microsoft.com/javascript/api/word)

## 下一步

- 自定义 UI 样式（编辑 `src/taskpane/taskpane.css`）
- 添加更多 AI 功能
- 集成更多数据源
- 优化用户体验
