/*
 * AI Copilot - Word 智能写作助手
 */

/* global console, document, Word, Office */

import aiService from '../services/aiService.js';
import mammoth from 'mammoth';

// 存储 AI 生成的内容
let generatedContent = "";

// 当前选中的角色
let currentRole = "default";

// 角色配置
const ROLE_CONFIGS = {
  default: {
    name: "通用助手",
    description: "通用 AI 助手，可以帮助你完成各种写作任务",
    systemPrompt: "你是一个专业的写作助手，能够帮助用户分析、生成和优化各类文档内容。"
  },
  storyteller: {
    name: "讲故事的人",
    description: "一款创意型人工智能聊天机器人，旨在通过想象力丰富的故事和传说吸引观众",
    systemPrompt: "你是一个富有创意的故事讲述者，擅长创作引人入胜的故事。你的故事充满想象力，情节生动，能够深深吸引读者。请用生动的语言、丰富的细节和引人入胜的情节来讲述故事。"
  },
  writer: {
    name: "文章润色",
    description: "一个帮助你轻松写出美丽而有创意的诗歌的工具",
    systemPrompt: "你是一个专业的文章润色专家，擅长优化文章的表达、结构和用词。你能够让文章更加流畅、优雅、专业，同时保持原意不变。"
  },
  novelist: {
    name: "小说家",
    description: "一个AI聊天机器人，作为虚拟小说家，创作出各种类型的引人入胜的故事",
    systemPrompt: "你是一个经验丰富的小说家，擅长创作各种类型的小说。你的作品情节跌宕起伏，人物形象鲜明，语言生动优美，能够深深打动读者的心。"
  },
  poet: {
    name: "诗人",
    description: "一个帮助创作优美诗歌的 AI 助手",
    systemPrompt: "你是一位才华横溢的诗人，擅长创作各种风格的诗歌。你的诗歌意境深远，韵律优美，富有感染力，能够触动人心。"
  },
  teacher: {
    name: "英语教官",
    description: "这个AI聊天机器人可以充当英语翻译、拼写纠正和改进的角色",
    systemPrompt: "你是一位专业的英语教师，擅长英语翻译、语法纠正和表达优化。你能够准确地翻译内容，纠正语法错误，并提供更地道的英语表达建议。"
  },
  programmer: {
    name: "程序员",
    description: "帮你编写和优化代码的 AI 助手",
    systemPrompt: "你是一位经验丰富的程序员，精通多种编程语言和开发技术。你能够编写高质量的代码，提供技术解决方案，并解释复杂的技术概念。"
  },
  marketer: {
    name: "营销专家",
    description: "帮助你创作营销文案和策划方案",
    systemPrompt: "你是一位资深的营销专家，擅长创作吸引人的营销文案、策划创意营销活动。你深谙消费者心理，能够创作出打动人心的营销内容。"
  },
  consultant: {
    name: "旅游顾问",
    description: "帮你规划旅行，让旅途充满惊喜和收获",
    systemPrompt: "你是一位经验丰富的旅游顾问，熟悉世界各地的旅游景点、文化习俗和旅行技巧。你能够为用户提供专业的旅行建议和详细的行程规划。"
  },
  twitter: {
    name: "推特文章创作",
    description: "帮你创作吸引眼球的推特内容，简洁有力，引发互动",
    systemPrompt: "你是一位专业的社交媒体内容创作者，擅长创作推特（Twitter/X）内容。你的文字简洁有力、富有感染力，善于使用话题标签、表情符号，能够引发用户互动和转发。你了解社交媒体传播规律，能够创作出病毒式传播的内容。"
  },
  crypto: {
    name: "币圈分析家",
    description: "专业的加密货币和区块链分析专家，提供深度市场洞察",
    systemPrompt: "你是一位资深的加密货币和区块链分析专家，对比特币、以太坊等主流加密货币以及 DeFi、NFT、Web3 等领域有深入研究。你能够分析市场趋势、解读技术指标、评估项目价值，并提供专业的投资建议。你的分析客观理性，基于数据和技术面，同时关注行业动态和监管政策。"
  }
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";
    
    // 绑定事件
    document.getElementById("analyzeBtn").onclick = analyzeSelectedText;
    document.getElementById("rewriteBtn").onclick = rewriteSelectedText;
    document.getElementById("checkPlagiarismBtn").onclick = checkPlagiarism;
    document.getElementById("reducePlagiarismBtn").onclick = reducePlagiarism;
    document.getElementById("formatPaperBtn").onclick = formatPaper;
    document.getElementById("imitateStyleBtn").onclick = imitateStyle;
    document.getElementById("generateBtn").onclick = generateContent;
    document.getElementById("uploadBtn").onclick = analyzeExcel;
    document.getElementById("insertBtn").onclick = insertToDocument;
    document.getElementById("roleSelector").onchange = handleRoleChange;
    
    // 文件上传显示文件名
    document.getElementById("sampleFile").onchange = function() {
      const file = this.files[0];
      if (file) {
        document.getElementById("fileName").textContent = file.name;
        document.getElementById("uploadStatus").classList.add("show");
      }
    };
    
    // 初始化角色描述
    updateRoleDescription();
  }
});

/**
 * 处理角色切换
 */
function handleRoleChange() {
  currentRole = document.getElementById("roleSelector").value;
  updateRoleDescription();
  showStatus(`已切换到：${ROLE_CONFIGS[currentRole].name}`, "success", 2000);
}

/**
 * 更新角色描述
 */
function updateRoleDescription() {
  const roleConfig = ROLE_CONFIGS[currentRole];
  document.getElementById("roleDescription").textContent = roleConfig.description;
}

/**
 * 显示加载动画
 */
function showLoading(message = "AI 正在生成中...") {
  const overlay = document.getElementById("loadingOverlay");
  const loadingText = document.getElementById("loadingText");
  loadingText.textContent = message;
  overlay.style.display = "flex";
}

/**
 * 隐藏加载动画
 */
function hideLoading() {
  const overlay = document.getElementById("loadingOverlay");
  overlay.style.display = "none";
}

/**
 * 显示状态消息
 */
function showStatus(message, type = "info", duration = 3000) {
  const statusMsg = document.getElementById("statusMsg");
  statusMsg.textContent = message;
  statusMsg.className = "status-msg show " + type;
  
  setTimeout(() => {
    statusMsg.classList.remove("show");
  }, duration);
}

/**
 * 显示结果区域
 */
function showResult(content) {
  generatedContent = content;
  document.getElementById("resultContent").textContent = content;
  document.getElementById("resultSection").style.display = "block";
}

/**
 * 1. 读取当前选中的文字（作为 AI 的上下文）
 */
async function getSelectedText() {
  return await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();
    return selection.text;
  });
}

/**
 * 2. 实时写入文字到 Word（优化格式，支持标题识别）
 */
async function writeContentToWord(text) {
  await Word.run(async (context) => {
    const body = context.document.body;
    
    // 先插入一个空行
    body.insertParagraph("", Word.InsertLocation.end);
    
    // 分段插入内容（按换行符分割）
    const paragraphs = text.split('\n');
    
    for (let i = 0; i < paragraphs.length; i++) {
      const paraText = paragraphs[i].trim();
      
      if (paraText) {
        // 插入有内容的段落
        const paragraph = body.insertParagraph(paraText, Word.InsertLocation.end);
        
        // 判断是否为标题（以【开头、包含"标题"、"##"开头、或全大写等）
        const isTitle = 
          paraText.startsWith('【') || 
          paraText.startsWith('##') ||
          paraText.startsWith('# ') ||
          (paraText.includes('标题') && paraText.length < 30) ||
          (paraText.endsWith('：') && paraText.length < 30);
        
        if (isTitle) {
          // 标题样式
          paragraph.font.size = 16; // 字号 16
          paragraph.font.name = "Microsoft YaHei"; // 微软雅黑
          paragraph.font.bold = true; // 加粗
          paragraph.font.color = "#0078d4"; // 蓝色
          paragraph.spaceAfter = 12; // 段后间距 12 磅
          paragraph.spaceBefore = 12; // 段前间距 12 磅
        } else {
          // 正文样式
          paragraph.font.size = 12; // 字号 12
          paragraph.font.name = "Microsoft YaHei"; // 微软雅黑
          paragraph.lineSpacing = 18; // 行间距 18 磅
          paragraph.spaceAfter = 6; // 段后间距 6 磅
        }
      } else {
        // 空行也插入，保持原有换行
        body.insertParagraph("", Word.InsertLocation.end);
      }
    }
    
    // 再插入一个空行
    body.insertParagraph("", Word.InsertLocation.end);
    
    await context.sync();
  });
}

/**
 * 3.7 查看查重率和 AI 率
 */
async function checkPlagiarism() {
  try {
    showStatus("正在读取选中文本...", "info");
    
    const selectedText = await getSelectedText();
    
    if (!selectedText || selectedText.trim() === "") {
      showStatus("请先选中文档中的文字！", "error");
      return;
    }
    
    showLoading("AI 正在检测查重率和 AI 率...");
    
    // 构建提示词
    const prompt = `请分析以下文本的原创性、重复度和 AI 生成痕迹。请按以下格式回复：

查重率：XX%
AI率：XX%
查重分析：[详细分析文本的原创性、常见表达、专业术语使用、与已有内容的相似度等]
AI率分析：[分析文本是否有 AI 生成的特征，如：过于规整的句式、机械化的表达、缺乏个性化语言等]
降重建议：[如何降低查重率的具体建议]
降AI建议：[如何让文本更像人类写作的建议]

文本内容：
${selectedText}`;
    
    // 调用 AI 服务
    const analysis = await aiService.generateContent(prompt);
    
    hideLoading();
    
    // 解析查重率和 AI 率
    const plagiarismMatch = analysis.match(/查重率[：:]\s*(\d+)%/);
    const aiMatch = analysis.match(/AI率[：:]\s*(\d+)%/);
    
    const plagiarismScore = plagiarismMatch ? parseInt(plagiarismMatch[1]) : 0;
    const aiScore = aiMatch ? parseInt(aiMatch[1]) : 0;
    
    // 显示结果
    document.getElementById("plagiarismScore").textContent = plagiarismScore + "%";
    document.getElementById("plagiarismScore").className = "score-value " + getScoreClass(plagiarismScore);
    
    document.getElementById("aiScore").textContent = aiScore + "%";
    document.getElementById("aiScore").className = "score-value " + getScoreClass(aiScore);
    
    document.getElementById("plagiarismAdvice").textContent = analysis;
    document.getElementById("plagiarismResult").style.display = "block";
    
    showStatus("✅ 检测完成！", "success");
    
  } catch (error) {
    console.error(error);
    hideLoading();
    showStatus("❌ 检测失败：" + error.message, "error");
  }
}

/**
 * 3.8 一键降重（同时降低查重率和 AI 率）
 */
async function reducePlagiarism() {
  try {
    showStatus("正在读取选中文本...", "info");
    
    const selectedText = await getSelectedText();
    
    if (!selectedText || selectedText.trim() === "") {
      showStatus("请先选中文档中的文字！", "error");
      return;
    }
    
    // 获取目标指标
    const targetPlagiarism = document.getElementById("targetPlagiarismRate").value;
    const targetAI = document.getElementById("targetAIRate").value;
    
    showLoading(`AI 正在降重中（目标：查重率${targetPlagiarism}%，AI率${targetAI}%）...`);
    
    // 获取当前角色的系统提示词
    const roleConfig = ROLE_CONFIGS[currentRole];
    
    // 构建提示词
    const prompt = `${roleConfig.systemPrompt}

请对以下文本进行深度改写，目标是将查重率降至 ${targetPlagiarism}% 以下，AI 率降至 ${targetAI}% 以下。

改写要求：

【降低查重率到 ${targetPlagiarism}%】
1. 保持原文的核心意思和观点完全不变
2. 大量使用同义词、近义词替换
3. 彻底改变句式结构和语序
4. 避免使用任何常见的固定表达
5. 增加具体的细节、数据和例子
6. 用自己的话重新组织内容

【降低 AI 率到 ${targetAI}%】
1. 使用非常自然、口语化的表达方式
2. 加入个人观点、情感色彩和主观判断
3. 避免过于规整、对称的句式结构
4. 使用俗语、比喻、拟人等多种修辞手法
5. 句子长短要有明显变化，不要整齐划一
6. 适当使用反问、设问等互动性表达
7. 加入一些口语化的连接词，如"其实"、"说实话"、"不过"等
8. 让文字有温度，像是在和朋友聊天

【重要提示】
- 查重率越低越好，最低可以到 5%
- AI 率越低越好，要让文字完全像人类写的
- 改写幅度要大，不要只是简单替换几个词
- 保持专业性的同时增加人性化表达

原文：
${selectedText}

请直接输出改写后的内容，不要包含任何解释说明。`;
    
    // 调用 AI 服务
    const rewrittenText = await aiService.generateContent(prompt);
    
    hideLoading();
    showResult(rewrittenText);
    showStatus(`✅ 降重完成！已按目标（查重率${targetPlagiarism}%，AI率${targetAI}%）改写，建议再次检测确认`, "success", 4000);
    
  } catch (error) {
    console.error(error);
    hideLoading();
    showStatus("❌ 降重失败：" + error.message, "error");
  }
}

/**
 * 根据查重率获取样式类
 */
function getScoreClass(score) {
  if (score < 20) return "score-good";
  if (score < 40) return "score-medium";
  return "score-high";
}

/**
 * 3.10 范文仿写
 */
async function imitateStyle() {
  try {
    // 获取范文文件
    const sampleFileInput = document.getElementById("sampleFile");
    const sampleFile = sampleFileInput.files[0];
    
    if (!sampleFile) {
      showStatus("请先上传范文文件！", "error");
      return;
    }
    
    showStatus("正在读取选中文本...", "info");
    
    // 获取当前选中的文本
    const selectedText = await getSelectedText();
    
    if (!selectedText || selectedText.trim() === "") {
      showStatus("请先选中文档中需要改写的文字！", "error");
      return;
    }
    
    showLoading("正在读取范文并分析风格...");
    
    // 读取范文内容
    const sampleContent = await readTextFile(sampleFile);
    
    // 获取仿写选项
    const imitateStructure = document.getElementById("imitateStructure").checked;
    const imitateStyle = document.getElementById("imitateStyle").checked;
    const imitateVocabulary = document.getElementById("imitateVocabulary").checked;
    const imitateTone = document.getElementById("imitateTone").checked;
    
    document.getElementById("loadingText").textContent = "AI 正在分析范文风格...";
    
    // 智能判断上传的是格式说明还是实际范文
    const isFormatGuide = sampleContent.includes("格式说明") || 
                          sampleContent.includes("版面設定") || 
                          sampleContent.includes("字型設定") ||
                          sampleContent.includes("論文架構") ||
                          sampleContent.includes("格式範例");
    
    let prompt = "";
    
    if (isFormatGuide) {
      // 如果是格式说明文档
      prompt = `你收到的是一份论文格式说明文档，而不是实际的范文内容。

【格式说明文档】
${sampleContent.substring(0, 4000)}

【我的文本内容】
${selectedText}

【任务要求】
请根据格式说明文档中的要求，将我的文本内容改写成符合该格式规范的论文。

具体要求：
1. 严格按照格式说明中的结构要求组织内容
2. 如果格式说明要求特定的章节（如：摘要、前言、研究方法、结果与讨论、结论等），请按此结构重新组织我的内容
3. 保持我的核心观点和内容不变，只调整结构和表达方式以符合格式要求
4. 如果格式说明中有字数限制（如摘要600字），请严格遵守
5. 如果格式说明要求特定的写作风格（如学术性、正式性），请调整语言风格
6. 添加格式说明中要求的必要元素（如关键字、计划编号等的占位符）

请直接输出改写后的内容，不要包含任何解释说明。`;
      
    } else {
      // 如果是实际范文
      prompt = `请仔细分析以下范文的特点，然后根据范文的风格改写我的文本。

【范文】
${sampleContent.substring(0, 3000)}

【需要改写的文本】
${selectedText}

【仿写要求】
`;

      if (imitateStructure) {
        prompt += `
1. 结构仿写：
   - 模仿范文的段落结构和层次
   - 学习范文的论证逻辑和展开方式
   - 参考范文的开头、过渡和结尾方式`;
      }
      
      if (imitateStyle) {
        prompt += `
2. 写作风格：
   - 模仿范文的句式特点（长短句搭配、句式变化等）
   - 学习范文的表达方式和修辞手法
   - 保持与范文相似的语言风格（正式/口语化、严谨/生动等）`;
      }
      
      if (imitateVocabulary) {
        prompt += `
3. 用词习惯：
   - 使用与范文相似的专业术语和词汇
   - 学习范文的用词精准度和丰富度
   - 模仿范文的词汇搭配习惯`;
      }
      
      if (imitateTone) {
        prompt += `
4. 语气语调：
   - 保持与范文一致的语气（客观/主观、严肃/轻松等）
   - 模仿范文的情感表达方式
   - 学习范文的态度和立场表达`;
      }
      
      prompt += `

【重要提示】
- 保持我的文本的核心内容和观点不变
- 只改变表达方式，使其更接近范文风格
- 确保改写后的内容自然流畅，不生硬
- 字数可以适当调整，但不要偏离太多

请直接输出改写后的内容，不要包含任何解释说明。`;
    }
    
    document.getElementById("loadingText").textContent = "AI 正在仿写中...";
    
    // 调用 AI 服务
    const rewrittenText = await aiService.generateContent(prompt);
    
    hideLoading();
    showResult(rewrittenText);
    
    if (isFormatGuide) {
      showStatus("✅ 已按格式要求改写完成！", "success");
    } else {
      showStatus("✅ 范文仿写完成！", "success");
    }
    
  } catch (error) {
    console.error(error);
    hideLoading();
    showStatus("❌ 仿写失败：" + error.message, "error");
  }
}

/**
 * 读取文本文件（支持 .txt、.doc、.docx）
 */
function readTextFile(file) {
  return new Promise((resolve, reject) => {
    const fileName = file.name.toLowerCase();
    
    if (fileName.endsWith('.txt')) {
      // 读取 TXT 文件
      const reader = new FileReader();
      reader.onload = (e) => {
        resolve(e.target.result);
      };
      reader.onerror = () => reject(new Error("TXT 文件读取失败"));
      reader.readAsText(file, 'UTF-8');
      
    } else if (fileName.endsWith('.docx') || fileName.endsWith('.doc')) {
      // 读取 Word 文件
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const arrayBuffer = e.target.result;
          const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
          const text = result.value;
          
          if (!text || text.trim().length === 0) {
            reject(new Error("Word 文档内容为空或无法读取"));
            return;
          }
          
          resolve(text);
        } catch (error) {
          reject(new Error("Word 文档解析失败：" + error.message));
        }
      };
      reader.onerror = () => reject(new Error("Word 文件读取失败"));
      reader.readAsArrayBuffer(file);
      
    } else {
      reject(new Error("不支持的文件格式，请上传 .txt、.doc 或 .docx 文件"));
    }
  });
}

/**
 * 3.9 论文格式化
 */
async function formatPaper() {
  try {
    showLoading("正在格式化论文...");
    
    // 获取格式设置
    const headerText = document.getElementById("headerText").value;
    const titleSize = parseInt(document.getElementById("titleSize").value);
    const heading1Size = parseInt(document.getElementById("heading1Size").value);
    const heading2Size = parseInt(document.getElementById("heading2Size").value);
    const bodySize = parseInt(document.getElementById("bodySize").value);
    const lineSpacing = parseFloat(document.getElementById("lineSpacing").value);
    
    await Word.run(async (context) => {
      const doc = context.document;
      
      // 1. 设置页眉
      if (headerText && headerText.trim()) {
        const sections = doc.sections;
        sections.load("items");
        await context.sync();
        
        for (let i = 0; i < sections.items.length; i++) {
          const header = sections.items[i].getHeader(Word.HeaderFooterType.primary);
          header.clear();
          const headerParagraph = header.insertParagraph(headerText, Word.InsertLocation.start);
          headerParagraph.alignment = Word.Alignment.centered;
          headerParagraph.font.size = 9;
          headerParagraph.font.name = "Microsoft YaHei";
        }
      }
      
      // 2. 格式化文档内容
      const body = doc.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();
      
      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load("text,style");
        await context.sync();
        
        const text = para.text.trim();
        if (!text) continue;
        
        // 判断段落类型并应用格式
        if (i === 0 || text.length < 50 && !text.includes("。") && !text.includes(".")) {
          // 可能是标题
          para.font.size = titleSize;
          para.font.bold = true;
          para.font.name = "Microsoft YaHei";
          para.alignment = Word.Alignment.centered;
          para.spaceAfter = 18;
          para.spaceBefore = 18;
        } else if (text.match(/^[一二三四五六七八九十\d]+[、\.．]/)) {
          // 一级标题（如：一、二、1. 2.）
          para.font.size = heading1Size;
          para.font.bold = true;
          para.font.name = "Microsoft YaHei";
          para.spaceAfter = 12;
          para.spaceBefore = 12;
        } else if (text.match(/^[\(（][一二三四五\d]+[\)）]/) || text.match(/^\d+\.\d+/)) {
          // 二级标题（如：（一）（1）1.1 1.2）
          para.font.size = heading2Size;
          para.font.bold = true;
          para.font.name = "Microsoft YaHei";
          para.spaceAfter = 6;
          para.spaceBefore = 6;
        } else {
          // 正文
          para.font.size = bodySize;
          para.font.name = "Microsoft YaHei";
          para.lineSpacing = lineSpacing * 12; // 转换为磅值
          para.spaceAfter = 0;
          para.firstLineIndent = 24; // 首行缩进2字符
        }
      }
      
      await context.sync();
    });
    
    hideLoading();
    showStatus("✅ 论文格式化完成！", "success");
    
  } catch (error) {
    console.error(error);
    hideLoading();
    showStatus("❌ 格式化失败：" + error.message, "error");
  }
}

/**
 * 3.5 扩写/缩写选中文本
 */
async function rewriteSelectedText() {
  try {
    showStatus("正在读取选中文本...", "info");
    
    const selectedText = await getSelectedText();
    
    if (!selectedText || selectedText.trim() === "") {
      showStatus("请先选中文档中的文字！", "error");
      return;
    }
    
    const rewriteType = document.getElementById("rewriteType").value;
    const targetWords = document.getElementById("targetWords").value;
    
    if (!targetWords || targetWords < 50) {
      showStatus("请输入目标字数（至少 50 字）！", "error");
      return;
    }
    
    const typeText = rewriteType === "expand" ? "扩写" : "缩写";
    showLoading(`AI 正在${typeText}中...`);
    
    // 获取当前角色的系统提示词
    const roleConfig = ROLE_CONFIGS[currentRole];
    
    // 构建提示词
    let prompt = `${roleConfig.systemPrompt}\n\n`;
    if (rewriteType === "expand") {
      prompt += `请将以下文本扩写到约 ${targetWords} 字，保持原意，增加细节、例子和描述，使内容更加丰富和生动：\n\n${selectedText}`;
    } else {
      prompt += `请将以下文本缩写到约 ${targetWords} 字，保留核心观点和关键信息，使表达更加简洁：\n\n${selectedText}`;
    }
    
    // 调用 AI 服务
    const rewrittenText = await aiService.generateContent(prompt);
    
    hideLoading();
    showResult(rewrittenText);
    showStatus(`✅ ${typeText}完成！`, "success");
    
  } catch (error) {
    console.error(error);
    hideLoading();
    showStatus("❌ 改写失败：" + error.message, "error");
  }
}

/**
 * 3. 分析选中文本
 */
async function analyzeSelectedText() {
  try {
    showStatus("正在读取选中文本...", "info");
    
    const selectedText = await getSelectedText();
    
    if (!selectedText || selectedText.trim() === "") {
      showStatus("请先选中文档中的文字！", "error");
      return;
    }
    
    showLoading("AI 正在分析中...");
    
    // 获取当前角色的系统提示词
    const roleConfig = ROLE_CONFIGS[currentRole];
    const customPrompt = `${roleConfig.systemPrompt}\n\n请分析以下文本，从结构、逻辑、用词、表达等多个维度给出专业的改进建议：\n\n${selectedText}`;
    
    // 调用 AI 服务分析文本
    const analysis = await aiService.generateContent(customPrompt);
    
    hideLoading();
    showResult(analysis);
    showStatus("✅ 分析完成！", "success");
    
  } catch (error) {
    console.error(error);
    hideLoading();
    showStatus("❌ 分析失败：" + error.message, "error");
  }
}

/**
 * 4. 生成内容
 */
async function generateContent() {
  try {
    const prompt = document.getElementById("promptInput").value;
    
    if (!prompt || prompt.trim() === "") {
      showStatus("请输入生成需求！", "error");
      return;
    }
    
    showLoading("AI 正在生成内容，请稍候...");
    
    // 获取当前角色的系统提示词
    const roleConfig = ROLE_CONFIGS[currentRole];
    const customPrompt = `${roleConfig.systemPrompt}\n\n${prompt}`;
    
    // 调用 AI 服务生成内容
    const content = await aiService.generateContent(customPrompt);
    
    hideLoading();
    showResult(content);
    showStatus("✅ 内容生成成功！", "success");
    
  } catch (error) {
    console.error(error);
    hideLoading();
    showStatus("❌ 生成失败：" + error.message, "error");
  }
}

/**
 * 5. 分析 Excel 文件
 */
async function analyzeExcel() {
  try {
    const fileInput = document.getElementById("excelInput");
    const file = fileInput.files[0];
    
    if (!file) {
      showStatus("请先选择 Excel 文件！", "error");
      return;
    }
    
    showLoading("正在读取 Excel 文件...");
    
    // 读取文件内容
    const fileData = await readExcelFile(file);
    
    document.getElementById("loadingText").textContent = "AI 正在分析数据...";
    
    // 调用 AI 服务分析 Excel 数据
    const analysis = await aiService.analyzeExcelData(fileData);
    
    hideLoading();
    showResult(analysis);
    showStatus("✅ Excel 分析完成！", "success");
    
  } catch (error) {
    console.error(error);
    hideLoading();
    showStatus("❌ 分析失败：" + error.message, "error");
  }
}

/**
 * 6. 插入内容到文档
 */
async function insertToDocument() {
  try {
    if (!generatedContent) {
      showStatus("没有可插入的内容！", "error");
      return;
    }
    
    showLoading("正在插入到文档...");
    await writeContentToWord(generatedContent);
    hideLoading();
    showStatus("✅ 内容已成功插入到文档！", "success");
    
  } catch (error) {
    console.error(error);
    hideLoading();
    showStatus("❌ 插入失败：" + error.message, "error");
  }
}

/**
 * 读取 Excel 文件
 */
function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      // 这里简化处理，实际需要使用 xlsx 库解析
      const data = e.target.result;
      resolve({
        name: file.name,
        size: file.size,
        data: data
      });
    };
    
    reader.onerror = () => reject(new Error("文件读取失败"));
    reader.readAsArrayBuffer(file);
  });
}

// AI 服务已通过 aiService 模块提供
// 如需自定义 AI 配置，请修改 src/config.js 文件

// 导出函数供测试使用
export { 
  getSelectedText, 
  writeContentToWord, 
  analyzeSelectedText,
  rewriteSelectedText,
  checkPlagiarism,
  reducePlagiarism,
  formatPaper,
  imitateStyle,
  generateContent,
  analyzeExcel
};
