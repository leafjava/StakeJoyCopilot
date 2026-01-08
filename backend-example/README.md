# AI Copilot 后端服务

这是一个简单的 Node.js 后端服务示例，用于处理 AI Copilot 插件的 AI 请求。

## 功能

- 文本分析（POST /api/analyze）
- 内容生成（POST /api/generate）
- Excel 数据分析（POST /api/excel-analysis）
- 健康检查（GET /health）

## 快速开始

### 1. 安装依赖

```bash
cd backend-example
npm install
```

### 2. 配置环境变量

复制 `.env.example` 为 `.env`：

```bash
copy .env.example .env
```

编辑 `.env` 文件，填入你的 OpenAI API Key：

```
OPENAI_API_KEY=sk-your-api-key-here
PORT=3001
```

### 3. 启动服务

```bash
npm start
```

或使用开发模式（自动重启）：

```bash
npm run dev
```

服务将运行在 `http://localhost:3001`

## API 文档

### 1. 文本分析

**请求**：
```http
POST /api/analyze
Content-Type: application/json

{
  "text": "要分析的文本内容"
}
```

**响应**：
```json
{
  "success": true,
  "analysis": "AI 分析结果...",
  "timestamp": "2024-01-08T10:00:00.000Z"
}
```

### 2. 内容生成

**请求**：
```http
POST /api/generate
Content-Type: application/json

{
  "prompt": "写一段关于人工智能的介绍"
}
```

**响应**：
```json
{
  "success": true,
  "content": "AI 生成的内容...",
  "timestamp": "2024-01-08T10:00:00.000Z"
}
```

### 3. Excel 数据分析

**请求**：
```http
POST /api/excel-analysis
Content-Type: application/json

{
  "name": "sales.xlsx",
  "data": [
    { "month": "Jan", "sales": 1000 },
    { "month": "Feb", "sales": 1200 }
  ]
}
```

**响应**：
```json
{
  "success": true,
  "analysis": "数据分析报告...",
  "timestamp": "2024-01-08T10:00:00.000Z"
}
```

## 配置前端

在前端项目的 `src/config.js` 中配置：

```javascript
export const AI_CONFIG = {
  provider: 'custom',
  custom: {
    baseURL: 'http://localhost:3001/api',
    endpoints: {
      analyze: '/analyze',
      generate: '/generate',
      excelAnalysis: '/excel-analysis'
    }
  }
};
```

## 部署到生产环境

### 使用 PM2

```bash
npm install -g pm2
pm2 start server.js --name ai-copilot-backend
pm2 save
pm2 startup
```

### 使用 Docker

创建 `Dockerfile`：

```dockerfile
FROM node:18-alpine
WORKDIR /app
COPY package*.json ./
RUN npm install --production
COPY . .
EXPOSE 3001
CMD ["node", "server.js"]
```

构建和运行：

```bash
docker build -t ai-copilot-backend .
docker run -p 3001:3001 --env-file .env ai-copilot-backend
```

## 安全建议

1. 使用 HTTPS（生产环境必须）
2. 实施 API 速率限制
3. 添加身份验证（JWT、API Key 等）
4. 验证和清理用户输入
5. 设置请求大小限制
6. 记录审计日志

## 扩展功能

- 添加缓存（Redis）减少 API 调用
- 实施请求队列处理高并发
- 添加用户认证和授权
- 集成多个 AI 模型
- 添加监控和日志系统
