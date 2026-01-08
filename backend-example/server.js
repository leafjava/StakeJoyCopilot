/**
 * AI Copilot 后端服务示例
 * 使用 Node.js + Express
 * 
 * 安装依赖：
 * npm install express cors dotenv openai
 */

const express = require('express');
const cors = require('cors');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3001;

// 中间件
app.use(cors());
app.use(express.json({ limit: '10mb' }));

// OpenAI 配置（可替换为其他 AI 服务）
const { Configuration, OpenAIApi } = require('openai');
const configuration = new Configuration({
  apiKey: process.env.OPENAI_API_KEY,
});
const openai = new OpenAIApi(configuration);

/**
 * 文本分析接口
 */
app.post('/api/analyze', async (req, res) => {
  try {
    const { text } = req.body;
    
    if (!text || text.trim() === '') {
      return res.status(400).json({ error: '文本不能为空' });
    }
    
    // 调用 OpenAI API
    const completion = await openai.createChatCompletion({
      model: 'gpt-4',
      messages: [
        {
          role: 'system',
          content: '你是一个专业的文本分析助手。请仔细分析用户提供的文本，从结构、逻辑、用词、表达等多个维度给出专业的改进建议。'
        },
        {
          role: 'user',
          content: text
        }
      ],
      max_tokens: 2000,
      temperature: 0.7
    });
    
    const analysis = completion.data.choices[0].message.content;
    
    res.json({
      success: true,
      analysis: analysis,
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    console.error('分析失败:', error);
    res.status(500).json({
      success: false,
      error: '分析失败，请稍后重试'
    });
  }
});

/**
 * 内容生成接口
 */
app.post('/api/generate', async (req, res) => {
  try {
    const { prompt } = req.body;
    
    if (!prompt || prompt.trim() === '') {
      return res.status(400).json({ error: '需求描述不能为空' });
    }
    
    const completion = await openai.createChatCompletion({
      model: 'gpt-4',
      messages: [
        {
          role: 'system',
          content: '你是一个专业的内容创作助手。请根据用户的需求生成高质量、结构清晰、逻辑严谨的内容。'
        },
        {
          role: 'user',
          content: prompt
        }
      ],
      max_tokens: 2000,
      temperature: 0.8
    });
    
    const content = completion.data.choices[0].message.content;
    
    res.json({
      success: true,
      content: content,
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    console.error('生成失败:', error);
    res.status(500).json({
      success: false,
      error: '生成失败，请稍后重试'
    });
  }
});

/**
 * Excel 数据分析接口
 */
app.post('/api/excel-analysis', async (req, res) => {
  try {
    const { data, name } = req.body;
    
    if (!data) {
      return res.status(400).json({ error: '数据不能为空' });
    }
    
    // 将数据转换为文本描述
    const dataDescription = JSON.stringify(data, null, 2);
    
    const completion = await openai.createChatCompletion({
      model: 'gpt-4',
      messages: [
        {
          role: 'system',
          content: '你是一个数据分析专家。请分析 Excel 数据，提取关键信息，发现数据趋势，并给出专业的分析报告和建议。'
        },
        {
          role: 'user',
          content: `请分析以下 Excel 数据（文件名：${name}）：\n\n${dataDescription}`
        }
      ],
      max_tokens: 2000,
      temperature: 0.7
    });
    
    const analysis = completion.data.choices[0].message.content;
    
    res.json({
      success: true,
      analysis: analysis,
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    console.error('分析失败:', error);
    res.status(500).json({
      success: false,
      error: '分析失败，请稍后重试'
    });
  }
});

/**
 * 健康检查接口
 */
app.get('/health', (req, res) => {
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString()
  });
});

// 启动服务器
app.listen(PORT, () => {
  console.log(`🚀 AI Copilot 后端服务运行在 http://localhost:${PORT}`);
  console.log(`📝 API 端点：`);
  console.log(`   - POST /api/analyze - 文本分析`);
  console.log(`   - POST /api/generate - 内容生成`);
  console.log(`   - POST /api/excel-analysis - Excel 分析`);
  console.log(`   - GET /health - 健康检查`);
});
