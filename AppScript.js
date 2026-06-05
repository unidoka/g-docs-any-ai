const GEMINI_MODEL = "gemini-3.5-flash";

function getConversationHistory() {
  const props = PropertiesService.getUserProperties();
  const history = props.getProperty('chatHistory');
  return history ? JSON.parse(history) : [];
}

function saveConversationHistory(history) {
  const props = PropertiesService.getUserProperties();
  props.setProperty('chatHistory', JSON.stringify(history));
}

function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('Gemini Docs')
      .addItem('Открыть чат', 'showSidebar')
      .addSeparator()
      .addToUi();
}

function showSidebar() {
  const htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 10px; margin: 0; }
    #chat { 
      height: calc(100vh - 120px); 
      overflow-y: auto; 
      border: 1px solid #ddd; 
      padding: 10px; 
      border-radius: 8px; 
      background: #f9f9f9;
      margin-bottom: 10px;
    }
    .message { 
      margin: 8px 0; 
      padding: 10px; 
      border-radius: 8px; 
      max-width: 85%; 
      word-wrap: break-word;
    }
    .user { 
      background: #e3f2fd; 
      margin-left: auto; 
      text-align: right; 
    }
    .ai { 
      background: #f1f8e9; 
      margin-right: auto; 
    }
    .input-container {
      display: flex;
      gap: 5px;
    }
    input { 
      flex: 1;
      padding: 10px; 
      border: 1px solid #ccc; 
      border-radius: 6px; 
    }
    button { 
      padding: 10px 15px;
      border: none;
      border-radius: 6px;
      background: #4CAF50;
      color: white;
      cursor: pointer;
    }
    button:hover { background: #45a049; }
    .insert-btn {
      margin-top: 8px;
      padding: 5px 10px;
      font-size: 11px;
      background: #2196F3;
    }
    .insert-btn:hover { background: #0b7dda; }
    .loading {
      text-align: center;
      color: #666;
      font-style: italic;
    }
  </style>
</head>
<body>
  <div id="chat"></div>
  
  <div class="input-container">
    <input type="text" id="userInput" placeholder="Напишите промпт для документации..." onkeypress="if(event.key==='Enter') sendMessage()">
    <button onclick="sendMessage()">Отправить</button>
  </div>

  <script>
    window.onload = function() {
      addMessage("Привет! Я помогу вам создать документацию. Опишите, что нужно написать.", false, false);
    };

    function addMessage(text, isUser, showInsertBtn = true) {
      const chat = document.getElementById('chat');
      const div = document.createElement('div');
      div.className = \`message \${isUser ? 'user' : 'ai'}\`;
      
      if (isUser || !showInsertBtn) {
        div.innerHTML = text;
      } else {
        div.innerHTML = text + \`<br><button class="insert-btn" onclick="insertResponse(this)">📄 Вставить в документ</button>\`;
      }
      
      chat.appendChild(div);
      chat.scrollTop = chat.scrollHeight;
    }

    function showLoading() {
      const chat = document.getElementById('chat');
      const div = document.createElement('div');
      div.className = 'message ai loading';
      div.id = 'loading-message';
      div.innerHTML = 'Gemini генерирует ответ...';
      chat.appendChild(div);
      chat.scrollTop = chat.scrollHeight;
    }

    function hideLoading() {
      const loading = document.getElementById('loading-message');
      if (loading) loading.remove();
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

    function insertResponse(btn) {
      const messageDiv = btn.parentElement;
      const messageText = messageDiv.innerText.replace('Вставить в документ', '').trim();
      
      btn.disabled = true;
      btn.innerText = 'Вставка...';
      
      google.script.run
        .withSuccessHandler((res) => {
          if (res.success) {
            btn.innerText = 'Вставлено!';
            btn.style.background = '#4CAF50';
          } else {
            btn.innerText = 'Ошибка';
            btn.style.background = '#f44336';
          }
        })
        .withFailureHandler((err) => {
          btn.innerText = 'Ошибка';
          btn.style.background = '#f44336';
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
      .setWidth(350);
  DocumentApp.getUi().showSidebar(html);
}

function sendChatMessage(userMessage) {
  try {
    const history = getConversationHistory();
    
    history.push({
      role: 'user',
      parts: [{ text: userMessage }]
    });

    const systemPrompt = `Ты - помощник по созданию технической документации. 
    Отвечай в формате Markdown. Используй:
    - # для заголовков
    - ## для подзаголовков
    - **жирный** для важных терминов
    - *курсив* для акцентов
    - Списки (- или 1.) для структурирования
    Отвечай чётко, по делу, на русском языке.`;

    const fullPrompt = systemPrompt + '\n\nЗапрос пользователя: ' + userMessage;
    
    const aiResponse = callGemini(fullPrompt);
    
    history.push({
      role: 'model',
      parts: [{ text: aiResponse }]
    });
    saveConversationHistory(history);
    
    return { success: true, response: aiResponse };
  } catch (e) {
    return { 
      success: false, 
      error: e.toString(),
      stack: e.stack 
    };
  }
}

function callGemini(prompt) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    throw new Error("API-ключ не найден! Добавьте GEMINI_API_KEY в Свойства скрипта.");
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`;
  
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { 
      temperature: 0.7, 
      maxOutputTokens: 8192 
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  });

  const json = JSON.parse(response.getContentText());
  return json.candidates[0].content.parts[0].text;
}

function insertMarkdownToDocument(markdown) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const result = insertMarkdownAtCursor(doc, markdown);
    
    if (result.success) {
      return { success: true, message: 'Текст успешно вставлен в документ!' };
    } else {
      return { success: false, message: result.error };
    }
  } catch (e) {
    return { success: false, message: 'Ошибка вставки: ' + e.toString() };
  }
}

function insertMarkdownFromCell() {
  const markdown = getMarkdownText();
  const doc = DocumentApp.getActiveDocument();
  insertMarkdownAtCursor(doc, markdown);
}

function getMarkdownText() {
  return "# Заголовок\n\nОбычный текст с **жирным** и *курсивом*.\n\n## Подзаголовок\n\n- Пункт 1\n- Пункт 2\n\n1. Первый\n2. Второй";
}

function insertMarkdownAtCursor(doc, markdown) {
  if (!doc) {
    doc = DocumentApp.getActiveDocument();
  }
  
  if (!markdown || typeof markdown !== 'string') {
    Logger.log('Ошибка: Текст Markdown не найден или пуст.');
    return { success: false, error: 'Текст Markdown не найден или пуст' };
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
    Logger.log('Ошибка вставки: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function applyInlineFormatting(paragraph) {
  const text = paragraph.getText();
  const textObj = paragraph.editAsText();
  
  const boldRegex = /\*\*(.+?)\*\*/g;
  let match;
  let processedText = text;
  const boldRanges = [];
  
  while ((match = boldRegex.exec(text)) !== null) {
    boldRanges.push({ start: match.index, end: match.index + match[0].length, content: match[1] });
  }
  
  for (let i = boldRanges.length - 1; i >= 0; i--) {
    const range = boldRanges[i];
    processedText = processedText.substring(0, range.start) + range.content + processedText.substring(range.end);
  }
  
  paragraph.setText(processedText);
  
  let offset = 0;
  for (const range of boldRanges) {
    const startPos = range.start - offset;
    const endPos = startPos + range.content.length - 1;
    textObj.setBold(startPos, endPos, true);
    offset += (range.end - range.start) - range.content.length;
  }
  
  const currentText = paragraph.getText();
  const italicRegex = /\*([^*]+?)\*/g;
  const italicRanges = [];
  
  while ((match = italicRegex.exec(currentText)) !== null) {
    italicRanges.push({ start: match.index, end: match.index + match[0].length, content: match[1] });
  }
  
  let italicProcessed = currentText;
  for (let i = italicRanges.length - 1; i >= 0; i--) {
    const range = italicRanges[i];
    italicProcessed = italicProcessed.substring(0, range.start) + range.content + italicProcessed.substring(range.end);
  }
  
  paragraph.setText(italicProcessed);
  
  offset = 0;
  for (const range of italicRanges) {
    const startPos = range.start - offset;
    const endPos = startPos + range.content.length - 1;
    textObj.setItalic(startPos, endPos, true);
    offset += (range.end - range.start) - range.content.length;
  }
}

function clearConversationHistory() {
  const props = PropertiesService.getUserProperties();
  props.deleteProperty('chatHistory');
}
