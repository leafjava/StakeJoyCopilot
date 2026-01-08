/**
 * AI API 配置文件示例
 * 
 * 使用方法：
 * 1. 复制此文件为 config.js
 * 2. 填入你的 API 配置信息
 * 3. 在 taskpane.js 中导入使用
 * 
 * 注意：config.js 应该添加到 .gitignore 中，避免泄露密钥
 */

export const AI_CONFIG = {
  // 选择 AI 服务提供商: 'openai' | 'azure' | 'custom'
  provider: 'openai',
  
  // OpenAI 配置
  openai: {
    apiKey: 'YOUR_OPENAI_API_KEY',
    baseURL: 'https://api.openai.com/v1',
    model: 'gpt-4',
    maxTokens: 2000,
    temperature: 0.7
  },
  
  // Azure OpenAI 配置
  azure: {
    endpoint: 'YOUR_AZURE_ENDPOINT',
    apiKey: 'YOUR_AZURE_API_KEY',
    deploymentName: 'YOUR_DEPLOYMENT_NAME',
    apiVersion: '2024-02-15-preview'
  },
  
  // 自定义后端配置
  custom: {
    baseURL: 'https://your-backend.com/api',
    apiKey: 'YOUR_CUSTOM_API_KEY',
    endpoints: {
      analyze: '/analyze',
      generate: '/generate',
      excelAnalysis: '/excel-analysis'
    }
  }
};

// 系统提示词配置
export const SYSTEM_PROMPTS = {
  analyze: '你是一个专业的文本分析助手。请仔细分析用户提供的文本，从结构、逻辑、用词、表达等多个维度给出专业的改进建议。',
  
  generate: '你是一个专业的内容创作助手。请根据用户的需求生成高质量、结构清晰、逻辑严谨的内容。',
  
  excelAnalysis: '你是一个数据分析专家。请分析 Excel 数据，提取关键信息，发现数据趋势，并给出专业的分析报告和建议。'
};

// 请求配置
export const REQUEST_CONFIG = {
  timeout: 30000, // 请求超时时间（毫秒）
  retryTimes: 3,  // 失败重试次数
  retryDelay: 1000 // 重试延迟（毫秒）
};
