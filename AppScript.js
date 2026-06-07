const GEMINI_MODEL = "gemini-3.5-flash";

function onOpen() {
  try {
    const ui = DocumentApp.getUi();
    ui.createMenu('Gemini Docs')
        .addItem('Открыть чат-ассистент', 'showSidebar')
        .addToUi();
  } catch (e) {
    Logger.log('onOpen() вызвана из неправильного контекста: ' + e.toString());
  }
}

function showSidebar() {
  const htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Google Sans', Arial, sans-serif; padding: 15px; margin: 0; color: #202124; }
    #chat { height: calc(100vh - 140px); overflow-y: auto; border: 1px solid #dadce0; padding: 12px; border-radius: 8px; background: #f8f9fa; margin-bottom: 12px; }
    .message { margin: 8px 0; padding: 10px 14px; border-radius: 12px; max-width: 90%; word-wrap: break-word; line-height: 1.4; font-size: 14px; }
    .user { background: #e8f0fe; margin-left: auto; text-align: right; color: #1967d2; }
    .ai { background: #e6f4ea; margin-right: auto; color: #137333; white-space: pre-wrap; }
    .input-container { display: flex; gap: 8px; }
    input { flex: 1; padding: 10px; border: 1px solid #dadce0; border-radius: 20px; outline: none; font-size: 14px; }
    input:focus { border-color: #1a73e8; }
    button { padding: 10px 16px; border: none; border-radius: 20px; background: #1a73e8; color: white; cursor: pointer; font-weight: 500; font-size: 14px; transition: background 0.2s; }
    button:hover { background: #1557b0; }
    button:disabled { background: #dadce0; color: #5f6368; cursor: not-allowed; }
    .insert-btn { margin-top: 8px; padding: 6px 12px; font-size: 12px; background: #1a73e8; border-radius: 4px; }
    .insert-btn:hover { background: #1557b0; }
    .loading { text-align: center; color: #5f6368; font-style: italic; padding: 10px; }
  </style>
</head>
<body>
  <div id="chat"></div>
  <div class="input-container">
    <input type="text" id="userInput" placeholder="Что добавить в документ?" onkeypress="if(event.key==='Enter') sendMessage()">
    <button id="sendBtn" onclick="sendMessage()">Отправить</button>
  </div>

  <script>
    window.onload = function() {
      addMessage("Привет! Я изучу стиль всего вашего документа и помогу написать новый текст в том же формате. Просто напишите, что нужно добавить.", false, false);
    };

    function addMessage(text, isUser, showInsertBtn = true) {
      const chat = document.getElementById('chat');
      const div = document.createElement('div');
      div.className = 'message ' + (isUser ? 'user' : 'ai');
      
      if (isUser || !showInsertBtn) {
        div.innerText = text;
      } else {
        div.innerText = text;
        const btn = document.createElement('button');
        btn.className = 'insert-btn';
        btn.innerText = 'Вставить в документ с форматированием';
        btn.onclick = function() { insertResponse(btn, text); };
        div.appendChild(document.createElement('br'));
        div.appendChild(btn);
      }
      
      chat.appendChild(div);
      chat.scrollTop = chat.scrollHeight;
    }

    function showLoading() {
      const chat = document.getElementById('chat');
      const div = document.createElement('div');
      div.className = 'message ai loading';
      div.id = 'loading-message';
      div.innerText = 'Анализирую стиль документа и генерирую ответ...';
      chat.appendChild(div);
      chat.scrollTop = chat.scrollHeight;
      document.getElementById('sendBtn').disabled = true;
    }

    function hideLoading() {
      const loading = document.getElementById('loading-message');
      if (loading) loading.remove();
      document.getElementById('sendBtn').disabled = false;
    }

    function sendMessage() {
      const input = document.getElementById('userInput');
      const text = input.value.trim();
      if (!text) return;

      addMessage(text, true);
      input.value = '';
      showLoading();

      google.script.run
        .withSuccessHandler((res) => {
          hideLoading();
          if (res.success) {
            addMessage(res.response, false);
          } else {
            addMessage("Ошибка: " + res.error, false, false);
          }
        })
        .withFailureHandler((err) => {
          hideLoading();
          addMessage("Критическая ошибка: " + err.message, false, false);
        })
        .sendChatMessage(text);
    }

    function insertResponse(btn, messageText) {
      btn.disabled = true;
      btn.innerText = 'Применяю форматирование и вставляю...';
      
      google.script.run
        .withSuccessHandler((res) => {
          if (res.success) {
            btn.innerText = 'Успешно вставлено!';
            btn.style.background = '#137333';
          } else {
            btn.innerText = 'Ошибка вставки';
            btn.style.background = '#d93025';
          }
        })
        .withFailureHandler((err) => {
          btn.innerText = 'Ошибка';
          btn.style.background = '#d93025';
          alert("Не удалось вставить: " + err.message);
        })
        .insertMarkdownToDocument(messageText);
    }
  </script>
</body>
</html>
  `;
  
  const html = HtmlService.createHtmlOutput(htmlContent)
      .setTitle('Gemini Documentation Assistant')
      .setWidth(400);
  DocumentApp.getUi().showSidebar(html);
}

function sendChatMessage(userMessage) {
  try {
    const docMarkdown = getDocumentAsMarkdown();
    
    const systemPrompt = `Ты - эксперт по технической документации. 
Ниже представлен ТЕКУЩИЙ ДОКУМЕНТ в формате Markdown. Внимательно проанализируй его структуру и стиль форматирования (уровни заголовков, использование списков, стиль изложения).

ТЕКУЩИЙ ДОКУМЕНТ:
\`\`\`markdown
${docMarkdown}
\`\`\`

ЗАДАЧА:
На основе проанализированного стиля этого документа, выполни запрос пользователя. 
Отвечай СТРОГО в формате Markdown, имитируя стиль исходного документа. Используй:
- #, ##, ### для заголовков соответствующих уровней
- **жирный** для ключевых терминов
- *курсив* для акцентов
- - или 1. для списков
Отвечай чётко, по делу, на русском языке. Не добавляй лишних комментариев вне markdown.`;

    const fullPrompt = systemPrompt + "\n\nЗАПРОС ПОЛЬЗОВАТЕЛЯ: " + userMessage;
    
    const aiResponse = callGemini(fullPrompt);
    
    return { success: true, response: aiResponse };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function callGemini(prompt) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    throw new Error("API-ключ не найден! Добавьте 'GEMINI_API_KEY' в Свойства скрипта (Project Settings -> Script Properties).");
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`;
  
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { 
      temperature: 0.4,
      maxOutputTokens: 8192 
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(response.getContentText());
  
  if (json.error) {
    throw new Error("Ошибка Gemini API: " + json.error.message);
  }
  
  if (!json.candidates || !json.candidates[0] || !json.candidates[0].content) {
    throw new Error("Пустой или некорректный ответ от Gemini");
  }
  
  return json.candidates[0].content.parts[0].text;
}

function getDocumentAsMarkdown() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  let markdown = "";
  const numChildren = body.getNumChildren();
  
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    const type = child.getType();
    
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const p = child.asParagraph();
      const text = p.getText().trim();
      if (!text) { 
        markdown += "\n"; 
        continue; 
      }
      
      const heading = p.getHeading();
      if (heading === DocumentApp.ParagraphHeading.HEADING1) {
        markdown += "# " + text + "\n\n";
      } else if (heading === DocumentApp.ParagraphHeading.HEADING2) {
        markdown += "## " + text + "\n\n";
      } else if (heading === DocumentApp.ParagraphHeading.HEADING3) {
        markdown += "### " + text + "\n\n";
      } else {
        markdown += text + "\n\n";
      }
    } 
    else if (type === DocumentApp.ElementType.LIST_ITEM) {
      const item = child.asListItem();
      const text = item.getText().trim();
      const glyph = item.getGlyphType();
      if (glyph === DocumentApp.GlyphType.BULLET) {
        markdown += "- " + text + "\n";
      } else {
        markdown += "1. " + text + "\n";
      }
    }
  }
  
  return markdown || "Документ пуст.";
}

function insertMarkdownToDocument(markdown) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const result = insertMarkdownAtCursor(doc, markdown);
    
    if (result.success) {
      return { success: true, message: 'Текст успешно вставлен с форматированием!' };
    } else {
      return { success: false, message: result.error };
    }
  } catch (e) {
    return { success: false, message: 'Ошибка вставки: ' + e.toString() };
  }
}

function insertMarkdownAtCursor(doc, markdown) {
  if (!markdown || typeof markdown !== 'string') {
    return { success: false, error: 'Текст для вставки пуст' };
  }

  try {
    const body = doc.getBody();
    const lines = markdown.split('\n');
    
    const cursor = doc.getCursor();
    let insertIndex = 0;

    if (cursor) {
      const element = cursor.getElement();
      let currentElement = element;

      while (currentElement.getParent() && 
             currentElement.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) {
        currentElement = currentElement.getParent();
      }
      
      insertIndex = body.getChildIndex(currentElement);
      
      if (element.getType() === DocumentApp.ElementType.TEXT) {
        const offset = cursor.getOffset();
        const fullText = element.asText().getText();
        
        if (offset > 0 && offset < fullText.length) {
          const before = fullText.substring(0, offset);
          const after = fullText.substring(offset);
          element.asText().setText(before);
          body.insertParagraph(insertIndex + 1, after);
          insertIndex++;
        } else if (offset === fullText.length) {
          insertIndex++;
        }
      } else {
        insertIndex++;
      }
    }

    let currentIndex = insertIndex;

    for (let line of lines) {
      const trimmed = line.trim();
      
      if (trimmed === '') {
        body.insertParagraph(currentIndex++, '');
        continue;
      }

      if (trimmed.startsWith('# ')) {
        body.insertParagraph(currentIndex++, trimmed.substring(2))
            .setHeading(DocumentApp.ParagraphHeading.HEADING1);
      } 
      else if (trimmed.startsWith('## ')) {
        body.insertParagraph(currentIndex++, trimmed.substring(3))
            .setHeading(DocumentApp.ParagraphHeading.HEADING2);
      } 
      else if (trimmed.startsWith('### ')) {
        body.insertParagraph(currentIndex++, trimmed.substring(4))
            .setHeading(DocumentApp.ParagraphHeading.HEADING3);
      } 
      else if (trimmed.match(/^\s*[-*+]\s/)) {
        const p = body.insertListItem(currentIndex++, trimmed.replace(/^\s*[-*+]\s/, ''));
        p.setGlyphType(DocumentApp.GlyphType.BULLET);
        applyInlineFormatting(p);
      } 
      else if (trimmed.match(/^\s*\d+\.\s/)) {
        const p = body.insertListItem(currentIndex++, trimmed.replace(/^\s*\d+\.\s/, ''));
        p.setGlyphType(DocumentApp.GlyphType.NUMBER);
        applyInlineFormatting(p);
      } 
      else {
        const p = body.insertParagraph(currentIndex++, trimmed);
        applyInlineFormatting(p);
      }
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function applyInlineFormatting(paragraph) {
  let text = paragraph.getText();
  const textObj = paragraph.editAsText();
  
  const boldRegex = /\*\*(.+?)\*\*/g;
  let match;
  const boldRanges = [];
  
  while ((match = boldRegex.exec(text)) !== null) {
    boldRanges.push({ start: match.index, end: match.index + match[0].length, content: match[1] });
  }
  
  let processedText = text;
  let offset = 0;
  
  for (let i = boldRanges.length - 1; i >= 0; i--) {
    const range = boldRanges[i];
    processedText = processedText.substring(0, range.start) + range.content + processedText.substring(range.end);
    
    const startPos = range.start - offset;
    const endPos = startPos + range.content.length - 1;
    textObj.setBold(startPos, endPos, true);
    
    offset += (range.end - range.start) - range.content.length;
  }
  
  text = processedText;
  const italicRegex = /\*([^*]+?)\*/g;
  const italicRanges = [];
  
  while ((match = italicRegex.exec(text)) !== null) {
    italicRanges.push({ start: match.index, end: match.index + match[0].length, content: match[1] });
  }
  
  processedText = text;
  offset = 0;
  
  for (let i = italicRanges.length - 1; i >= 0; i--) {
    const range = italicRanges[i];
    processedText = processedText.substring(0, range.start) + range.content + processedText.substring(range.end);
    
    const startPos = range.start - offset;
    const endPos = startPos + range.content.length - 1;
    textObj.setItalic(startPos, endPos, true);
    
    offset += (range.end - range.start) - range.content.length;
  }
  
  paragraph.setText(processedText);
}

function getConversationHistory() {
  return [];
}
