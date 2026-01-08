/**
 * AI 服务封装
 * 支持多种 AI 提供商的统一接口
 */

// 如果创建了 config.js，取消下面的注释
// import { AI_CONFIG, SYSTEM_PROMPTS, REQUEST_CONFIG } from '../config.js';

/**
 * AI 服务类
 */
class AIService {
  constructor(config) {
    this.config = config || {
      provider: 'mock', // 默认使用模拟数据
      timeout: 30000
    };
  }

  /**
   * 分析文本
   */
  async analyzeText(text) {
    switch (this.config.provider) {
      case 'openai':
        return await this.callOpenAI(text, 'analyze');
      case 'azure':
        return await this.callAzureOpenAI(text, 'analyze');
      case 'custom':
        return await this.callCustomAPI(text, 'analyze');
      default:
        return await this.mockAnalyze(text);
    }
  }

  /**
   * 生成内容
   */
  async generateContent(prompt) {
    switch (this.config.provider) {
      case 'openai':
        return await this.callOpenAI(prompt, 'generate');
      case 'azure':
        return await this.callAzureOpenAI(prompt, 'generate');
      case 'custom':
        return await this.callCustomAPI(prompt, 'generate');
      default:
        return await this.mockGenerate(prompt);
    }
  }

  /**
   * 分析 Excel 数据
   */
  async analyzeExcelData(data) {
    switch (this.config.provider) {
      case 'openai':
        return await this.callOpenAI(JSON.stringify(data), 'excelAnalysis');
      case 'azure':
        return await this.callAzureOpenAI(JSON.stringify(data), 'excelAnalysis');
      case 'custom':
        return await this.callCustomAPI(data, 'excelAnalysis');
      default:
        return await this.mockExcelAnalysis(data);
    }
  }

  /**
   * 调用 OpenAI API
   */
  async callOpenAI(content, type) {
    const { apiKey, baseURL, model, maxTokens, temperature } = this.config.openai;
    
    const systemPrompt = this.getSystemPrompt(type);
    
    const response = await fetch(`${baseURL}/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: model,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: content }
        ],
        max_tokens: maxTokens,
        temperature: temperature
      })
    });

    if (!response.ok) {
      throw new Error(`OpenAI API 错误: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    return data.choices[0].message.content;
  }

  /**
   * 调用 Azure OpenAI API
   */
  async callAzureOpenAI(content, type) {
    const { endpoint, apiKey, deploymentName, apiVersion } = this.config.azure;
    
    const systemPrompt = this.getSystemPrompt(type);
    
    const url = `${endpoint}/openai/deployments/${deploymentName}/chat/completions?api-version=${apiVersion}`;
    
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'api-key': apiKey
      },
      body: JSON.stringify({
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: content }
        ]
      })
    });

    if (!response.ok) {
      throw new Error(`Azure OpenAI API 错误: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    return data.choices[0].message.content;
  }

  /**
   * 调用自定义后端 API
   */
  async callCustomAPI(content, type) {
    const { baseURL, apiKey, endpoints } = this.config.custom;
    
    const endpoint = endpoints[type];
    const url = `${baseURL}${endpoint}`;
    
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        content: content,
        type: type
      })
    });

    if (!response.ok) {
      throw new Error(`自定义 API 错误: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    return data.result || data.content || data.analysis;
  }

  /**
   * 获取系统提示词
   */
  getSystemPrompt(type) {
    const prompts = {
      analyze: '你是一个专业的文本分析助手。请仔细分析用户提供的文本，从结构、逻辑、用词、表达等多个维度给出专业的改进建议。',
      generate: '你是一个专业的内容创作助手。请根据用户的需求生成高质量、结构清晰、逻辑严谨的内容。',
      excelAnalysis: '你是一个数据分析专家。请分析数据，提取关键信息，发现数据趋势，并给出专业的分析报告和建议。'
    };
    
    return prompts[type] || prompts.generate;
  }

  // ========== 模拟 API（用于测试） ==========

  async mockAnalyze(text) {
    await this.delay(1500);
    
    return `【AI 分析结果】

原文内容：
${text}

分析建议：
1. 文本结构清晰，逻辑连贯
2. 建议增加更多具体案例支撑观点
3. 可以优化部分用词，使表达更加专业
4. 建议添加数据支持，增强说服力

改进建议：
- 在关键论点后增加实例说明
- 使用更精准的专业术语
- 适当添加过渡句，提升流畅度

---
提示：这是模拟数据，请配置真实的 AI API。`;
  }

  async mockGenerate(prompt) {
    await this.delay(2000);
    
    return `【AI 生成内容】

根据您的需求"${prompt}"，生成以下内容：

人工智能（Artificial Intelligence，简称 AI）是计算机科学的一个重要分支，致力于研究、开发用于模拟、延伸和扩展人的智能的理论、方法、技术及应用系统。

AI 的核心目标是让机器能够像人类一样思考、学习和解决问题。近年来，随着深度学习、神经网络等技术的突破，AI 在图像识别、自然语言处理、自动驾驶等领域取得了显著进展。

未来，人工智能将继续深刻影响我们的生活和工作方式，为人类社会带来更多可能性。

---
提示：这是模拟数据，请配置真实的 AI API。`;
  }

  async mockExcelAnalysis(data) {
    await this.delay(2000);
    
    return `【Excel 数据分析报告】

文件名：${data.name || '未知'}
文件大小：${data.size ? (data.size / 1024).toFixed(2) + ' KB' : '未知'}

数据概览：
- 检测到多个工作表
- 数据行数：约 150 行
- 数据列数：8 列

关键发现：
1. 销售数据呈现上升趋势，同比增长 23%
2. 第二季度表现最佳，占全年销售额的 35%
3. 产品 A 系列贡献了 60% 的收入
4. 客户满意度平均分为 4.2/5.0

建议：
- 加大对高绩效产品的投入
- 优化第四季度的营销策略
- 关注客户反馈，持续改进服务质量

---
提示：这是模拟数据，请配置真实的 AI API。`;
  }

  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

// 创建默认实例
const aiService = new AIService();

export default aiService;
export { AIService };
