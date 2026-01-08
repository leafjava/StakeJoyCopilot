/*
 * AI Copilot - Word 智能写作助手
 */

/* global console, document, Word, Office */

import aiService from '../services/aiService.js';

// 存储 AI 生成的内容
let generatedContent = "";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";
    
    // 绑定事件
    document.getElementById("analyzeBtn").onclick = analyzeSelectedText;
    document.getElementById("rewriteBtn").onclick = rewriteSelectedText;
    document.getElementById("generateBtn").onclick = generateContent;
    document.getElementById("uploadBtn").onclick = analyzeExcel;
    document.getElementById("insertBtn").onclick = insertToDocument;
  }
});

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
    
    // 构建提示词
    let prompt = "";
    if (rewriteType === "expand") {
      prompt = `请将以下文本扩写到约 ${targetWords} 字，保持原意，增加细节、例子和描述，使内容更加丰富和生动：\n\n${selectedText}`;
    } else {
      prompt = `请将以下文本缩写到约 ${targetWords} 字，保留核心观点和关键信息，使表达更加简洁：\n\n${selectedText}`;
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
    
    // 调用 AI 服务分析文本
    const analysis = await aiService.analyzeText(selectedText);
    
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
    
    // 调用 AI 服务生成内容
    const content = await aiService.generateContent(prompt);
    
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
  generateContent,
  analyzeExcel
};
