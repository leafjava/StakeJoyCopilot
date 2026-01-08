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
    document.getElementById("generateBtn").onclick = generateContent;
    document.getElementById("uploadBtn").onclick = analyzeExcel;
    document.getElementById("insertBtn").onclick = insertToDocument;
  }
});

/**
 * 显示状态消息
 */
function showStatus(message, duration = 3000) {
  const statusMsg = document.getElementById("statusMsg");
  statusMsg.textContent = message;
  statusMsg.classList.add("show");
  
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
 * 2. 实时写入文字到 Word
 */
async function writeContentToWord(text) {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertParagraph(text, Word.InsertLocation.end);
    await context.sync();
  });
}

/**
 * 3. 分析选中文本
 */
async function analyzeSelectedText() {
  try {
    showStatus("正在读取选中文本...");
    
    const selectedText = await getSelectedText();
    
    if (!selectedText || selectedText.trim() === "") {
      showStatus("请先选中文档中的文字！");
      return;
    }
    
    showStatus("正在调用 AI 分析...");
    
    // 调用 AI 服务分析文本
    const analysis = await aiService.analyzeText(selectedText);
    
    showResult(analysis);
    showStatus("分析完成！");
    
  } catch (error) {
    console.error(error);
    showStatus("分析失败：" + error.message);
  }
}

/**
 * 4. 生成内容
 */
async function generateContent() {
  try {
    const prompt = document.getElementById("promptInput").value;
    
    if (!prompt || prompt.trim() === "") {
      showStatus("请输入生成需求！");
      return;
    }
    
    showStatus("AI 正在生成内容...");
    
    // 调用 AI 服务生成内容
    const content = await aiService.generateContent(prompt);
    
    showResult(content);
    showStatus("生成完成！");
    
  } catch (error) {
    console.error(error);
    showStatus("生成失败：" + error.message);
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
      showStatus("请先选择 Excel 文件！");
      return;
    }
    
    showStatus("正在读取 Excel 文件...");
    
    // 读取文件内容
    const fileData = await readExcelFile(file);
    
    showStatus("正在调用 AI 分析数据...");
    
    // 调用 AI 服务分析 Excel 数据
    const analysis = await aiService.analyzeExcelData(fileData);
    
    showResult(analysis);
    showStatus("Excel 分析完成！");
    
  } catch (error) {
    console.error(error);
    showStatus("分析失败：" + error.message);
  }
}

/**
 * 6. 插入内容到文档
 */
async function insertToDocument() {
  try {
    if (!generatedContent) {
      showStatus("没有可插入的内容！");
      return;
    }
    
    showStatus("正在插入到文档...");
    await writeContentToWord(generatedContent);
    showStatus("插入成功！");
    
  } catch (error) {
    console.error(error);
    showStatus("插入失败：" + error.message);
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
  generateContent,
  analyzeExcel
};
