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
    * { box-sizing: border-box; }
    body { 
      font-family: 'Google Sans', Arial, sans-serif; 
      padding: 10px; 
      margin: 0; 
      color: #202124;
      height: 100vh;
      display: flex;
      flex-direction: column;
      overflow: hidden;
    }
    #chat { 
      flex: 1;
      overflow-y: auto; 
      border: 1px solid #dadce0; 
      padding: 8px; 
      border-radius: 8px; 
      background: #f8f9fa;
      margin-bottom: 8px;
      min-height: 0;
    }
    .message { 
      margin: 6px 0; 
      padding: 8px 10px; 
      border-radius: 8px; 
      max-width: 100%; 
      word-wrap: break-word; 
      line-height: 1.3; 
      font-size: 13px; 
    }
    .user { 
      background: #e8f0fe; 
      margin-left: auto; 
      text-align: right; 
      color: #1967d2; 
    }
    .ai { 
      background: #e6f4ea; 
      margin-right: auto; 
      color: #137333; 
      white-space: pre-wrap; 
    }
    .controls { 
      display: flex;
      flex-direction: column;
      gap: 6px;
    }
    .options-row {
      display: flex;
      gap: 6px;
      align-items: center;
      font-size: 12px;
    }
    .options-row label {
      display: flex;
      align-items: center;
      gap: 4px;
      cursor: pointer;
    }
    .image-section { 
      display: flex; 
      gap: 6px; 
      align-items: center;
      flex-wrap: wrap;
    }
    input[type="file"] { display: none; }
    .file-label { 
      padding: 6px 12px; 
      border-radius: 16px; 
      background: #34a853; 
      color: white; 
      cursor: pointer; 
      font-size: 12px; 
      white-space: nowrap;
    }
    .file-label:hover { background: #2d8f47; }
    .image-preview { 
      max-width: 100px; 
      max-height: 50px; 
      border-radius: 4px; 
      display: none;
    }
    .input-container { 
      display: flex; 
      gap: 6px; 
    }
    input[type="text"] { 
      flex: 1; 
      padding: 8px 12px; 
      border: 1px solid #dadce0; 
      border-radius: 16px; 
      outline: none; 
      font-size: 13px; 
    }
    input[type="text"]:focus { border-color: #1a73e8; }
    button { 
      padding: 8px 14px; 
      border: none; 
      border-radius: 16px; 
      background: #1a73e8; 
      color: white; 
      cursor: pointer; 
      font-size: 13px; 
      white-space: nowrap;
    }
    button:hover { background: #1557b0; }
    button:disabled { background: #dadce0; cursor: not-allowed; }
    .insert-btn { 
      margin-top: 6px; 
      padding: 4px 10px; 
      font-size: 11px; 
      background: #1a73e8; 
      border-radius: 4px; 
    }
    .loading { 
      text-align: center; 
      color: #5f6368; 
      font-style: italic; 
      padding: 8px; 
      font-size: 12px;
    }
    #imageName { 
      font-size: 11px; 
      color: #5f6368;
      max-width: 150px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
  </style>
</head>
<body>
  <div id="chat"></div>
  
  <div class="controls">
    <div class="image-section">
      <label for="imageInput" class="file-label">загрузить изображение</label>
      <input type="file" id="imageInput" accept="image/*" onchange="previewImage()">
      <span id="imageName"></span>
      <img id="imagePreview" class="image-preview">
    </div>
    
    <div class="options-row">
      <label>
        <input type="radio" name="genType" value="paragraph" checked>
        Абзац
      </label>
      <label>
        <input type="radio" name="genType" value="sentence">
        Предложение
      </label>
    </div>
    
    <div class="input-container">
      <input type="text" id="userInput" placeholder="Что добавить?" onkeypress="if(event.key==='Enter') sendMessage()">
      <button id="sendBtn" onclick="sendMessage()">Отправить</button>
    </div>
  </div>

  <script>
    let selectedImageBase64 = null;
    let selectedImageName = '';

    window.onload = function() {
      addMessage("Привет! Опишите, что нужно добавить в текущий раздел. Можно прикрепить изображение.", false, false);
    };

    function previewImage() {
      const fileInput = document.getElementById('imageInput');
      const file = fileInput.files[0];
      
      if (!file) {
        selectedImageBase64 = null;
        selectedImageName = '';
        document.getElementById('imageName').innerText = '';
        document.getElementById('imagePreview').style.display = 'none';
        return;
      }

      const reader = new FileReader();
      reader.onload = function(e) {
        selectedImageBase64 = e.target.result.split(',')[1];
        selectedImageName = file.name;
        
        document.getElementById('imageName').innerText = file.name;
        document.getElementById('imagePreview').src = e.target.result;
        document.getElementById('imagePreview').style.display = 'block';
      };
      reader.readAsDataURL(file);
    }

    function addMessage(text, isUser, showInsertBtn) {
      const chat = document.getElementById('chat');
      const div = document.createElement('div');
      div.className = 'message ' + (isUser ? 'user' : 'ai');
      
      if (isUser || showInsertBtn === false) {
        div.innerText = text;
      } else {
        div.innerText = text;
        const btn = document.createElement('button');
        btn.className = 'insert-btn';
        btn.innerText = '📄 Вставить в документ';
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
      div.innerText = '⏳ Анализ раздела и генерация...';
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
      if (!text && !selectedImageBase64) return;

      const genType = document.querySelector('input[name="genType"]:checked').value;
      
      let messageText = text;
      if (selectedImageBase64) {
        messageText += (text ? '\\n\\n' : '') + '[Изображение: ' + selectedImageName + ']';
      }
      messageText += ' [' + (genType === 'paragraph' ? 'Абзац' : 'Предложение') + ']';

      addMessage(messageText, true);
      input.value = '';
      
      const imageToSend = selectedImageBase64;
      const imageNameToSend = selectedImageName;
      const hasImage = !!selectedImageBase64;
      
      if (hasImage) {
        clearImage();
      }
      
      showLoading();

      if (hasImage) {
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
            addMessage("Ошибка: " + err.message, false, false);
          })
          .sendChatMessageWithImage(text, imageToSend, imageNameToSend, genType);
      } else {
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
            addMessage("Ошибка: " + err.message, false, false);
          })
          .sendChatMessage(text, genType);
      }
    }

    function clearImage() {
      selectedImageBase64 = null;
      selectedImageName = '';
      document.getElementById('imageInput').value = '';
      document.getElementById('imageName').innerText = '';
      document.getElementById('imagePreview').style.display = 'none';
    }

    function insertResponse(btn, messageText) {
      btn.disabled = true;
      btn.innerText = '⏳ Вставка...';
      
      google.script.run
        .withSuccessHandler((res) => {
          if (res.success) {
            btn.innerText = '✅ Вставлено!';
            btn.style.background = '#137333';
          } else {
            btn.innerText = '❌ Ошибка';
            btn.style.background = '#d93025';
          }
        })
        .withFailureHandler((err) => {
          btn.innerText = '❌ Ошибка';
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
      .setWidth(350);
  DocumentApp.getUi().showSidebar(html);
}

function sendChatMessage(userMessage, genType) {
  try {
    // Получаем ТОЛЬКО текущий раздел где находится курсор
    const currentSection = getCurrentSectionMarkdown();
    
    // Получаем стили всего документа
    const docStyles = getDocumentStyleGuide();
    
    const genTypeText = genType === 'paragraph' ? 'абзац (3-5 предложений)' : '1-2 предложения';

    const systemPrompt = 'Ты - эксперт по технической документации.\n\n' +
    'СТИЛЬ ДОКУМЕНТА (копируй это форматирование):\n' +
    '```markdown\n' + docStyles + '\n```\n\n' +
    'ТЕКУЩИЙ РАЗДЕЛ (контекст):\n' +
    '```markdown\n' + currentSection + '\n```\n\n' +
    'ЗАДАЧА:\n' +
    'Сгенерируй ' + genTypeText + ' для текущего раздела на основе запроса пользователя. ' +
    'Строго следуй стилю документа. Отвечай ТОЛЬКО текстом без markdown-разметки (кроме **жирного** и *курсива*).';

    const fullPrompt = systemPrompt + "\n\nЗАПРОС: " + userMessage;
    
    const aiResponse = callGemini(fullPrompt);
    
    return { success: true, response: aiResponse };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function sendChatMessageWithImage(userMessage, imageBase64, imageName, genType) {
  try {
    const currentSection = getCurrentSectionMarkdown();
    const docStyles = getDocumentStyleGuide();
    
    const genTypeText = genType === 'paragraph' ? 'абзац (3-5 предложений)' : '1-2 предложения';

    const systemPrompt = 'Ты - эксперт по технической документации с vision capabilities.\n\n' +
    'СТИЛЬ ДОКУМЕНТА:\n' +
    '```markdown\n' + docStyles + '\n```\n\n' +
    'ТЕКУЩИЙ РАЗДЕЛ:\n' +
    '```markdown\n' + currentSection + '\n```\n\n' +
    'ЗАДАЧА:\n' +
    'Пользователь загрузил изображение "' + imageName + '" и просит: "' + userMessage + '"\n' +
    'Проанализируй изображение и сгенерируй ' + genTypeText + ' для текущего раздела. ' +
    'Следуй стилю документа. Отвечай ТОЛЬКО текстом.';

    // Отправляем текст + изображение в Gemini
    const aiResponse = callGeminiWithImage(systemPrompt, imageBase64);
    
    return { success: true, response: aiResponse };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Получает текущий раздел где находится курсор
function getCurrentSectionMarkdown() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const cursor = doc.getCursor();
  
  if (!cursor) {
    return getDocumentAsMarkdown();
  }
  
  const element = cursor.getElement();
  let currentElement = element;
  
  // Находим индекс текущего элемента
  while (currentElement.getParent() && 
         currentElement.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) {
    currentElement = currentElement.getParent();
  }
  
  const currentIndex = body.getChildIndex(currentElement);
  
  // Ищем ближайший заголовок выше
  let sectionStart = 0;
  let sectionHeading = '';
  
  for (let i = currentIndex; i >= 0; i--) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = child.asParagraph();
      const heading = p.getHeading();
      if (heading === DocumentApp.ParagraphHeading.HEADING1 || 
          heading === DocumentApp.ParagraphHeading.HEADING2 || 
          heading === DocumentApp.ParagraphHeading.HEADING3) {
        sectionStart = i;
        sectionHeading = p.getText();
        break;
      }
    }
  }
  
  // Ищем конец раздела (следующий заголовок того же или более высокого уровня)
  let sectionEnd = body.getNumChildren();
  let currentHeadingLevel = 0;
  
  if (sectionHeading) {
    const startPara = body.getChild(sectionStart).asParagraph();
    const startHeading = startPara.getHeading();
    if (startHeading === DocumentApp.ParagraphHeading.HEADING1) currentHeadingLevel = 1;
    else if (startHeading === DocumentApp.ParagraphHeading.HEADING2) currentHeadingLevel = 2;
    else if (startHeading === DocumentApp.ParagraphHeading.HEADING3) currentHeadingLevel = 3;
  }
  
  for (let i = currentIndex + 1; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = child.asParagraph();
      const heading = p.getHeading();
      let headingLevel = 0;
      if (heading === DocumentApp.ParagraphHeading.HEADING1) headingLevel = 1;
      else if (heading === DocumentApp.ParagraphHeading.HEADING2) headingLevel = 2;
      else if (heading === DocumentApp.ParagraphHeading.HEADING3) headingLevel = 3;
      
      if (headingLevel > 0 && headingLevel <= currentHeadingLevel) {
        sectionEnd = i;
        break;
      }
    }
  }
  
  // Собираем markdown для раздела
  let markdown = '';
  if (sectionHeading) {
    markdown += '# ' + sectionHeading + '\n\n';
  }
  
  for (let i = sectionStart + 1; i < sectionEnd && i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    const type = child.getType();
    
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const p = child.asParagraph();
      const text = p.getText().trim();
      if (text) {
        markdown += text + '\n\n';
      }
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      const item = child.asListItem();
      const text = item.getText().trim();
      markdown += '- ' + text + '\n';
    }
  }
  
  return markdown || "Раздел пуст.";
}

// Получает стили всего документа (заголовки, списки и т.д.)
function getDocumentStyleGuide() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  let styles = {
    headings: [],
    hasBullets: false,
    hasNumbers: false
  };
  
  const numChildren = body.getNumChildren();
  
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    const type = child.getType();
    
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const p = child.asParagraph();
      const text = p.getText();
      const heading = p.getHeading();
      
      if (heading === DocumentApp.ParagraphHeading.HEADING1) {
        styles.headings.push({ level: 1, text: text });
      } else if (heading === DocumentApp.ParagraphHeading.HEADING2) {
        styles.headings.push({ level: 2, text: text });
      } else if (heading === DocumentApp.ParagraphHeading.HEADING3) {
        styles.headings.push({ level: 3, text: text });
      }
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      const item = child.asListItem();
      if (item.getGlyphType() === DocumentApp.GlyphType.BULLET) {
        styles.hasBullets = true;
      } else {
        styles.hasNumbers = true;
      }
    }
  }
  
  // Формируем описание стиля
  let styleGuide = 'Структура документа:\n';
  styles.headings.slice(0, 10).forEach(h => {
    const prefix = '#'.repeat(h.level);
    styleGuide += prefix + ' ' + h.text + '\n';
  });
  
  styleGuide += '\nПравила форматирования:\n';
  styleGuide += '- Используй **жирный** для ключевых терминов\n';
  styleGuide += '- Используй *курсив* для акцентов\n';
  if (styles.hasBullets) styleGuide += '- Используй маркированные списки (-)\n';
  if (styles.hasNumbers) styleGuide += '- Используй нумерованные списки (1.)\n';
  
  return styleGuide;
}

function callGemini(prompt) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    throw new Error("API-ключ не найден!");
  }

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + apiKey;
  
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { 
      temperature: 0.4,
      maxOutputTokens: 2048 
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
    throw new Error("Пустой ответ от Gemini");
  }
  
  return json.candidates[0].content.parts[0].text;
}

function callGeminiWithImage(prompt, imageBase64) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    throw new Error("API-ключ не найден!");
  }

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + apiKey;
  
  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        {
          inlineData: {
            mimeType: 'image/png',
            data: imageBase64
          }
        }
      ]
    }],
    generationConfig: { 
      temperature: 0.4,
      maxOutputTokens: 2048 
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
    throw new Error("Пустой ответ от Gemini");
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
    else if (type === DocumentApp.ElementType.INLINE_IMAGE) {
      markdown += "[Изображение]\n\n";
    }
  }
  
  return markdown || "Документ пуст.";
}

function insertMarkdownToDocument(markdown) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const result = insertMarkdownAtCursor(doc, markdown);
    
    if (result.success) {
      return { success: true, message: 'Текст успешно вставлен!' };
    } else {
      return { success: false, message: result.error };
    }
  } catch (e) {
    return { success: false, message: 'Ошибка: ' + e.toString() };
  }
}

function insertMarkdownAtCursor(doc, markdown) {
  if (!markdown || typeof markdown !== 'string') {
    return { success: false, error: 'Текст пуст' };
  }

  try {
    const body = doc.getBody();
    const lines = markdown.split('\n');
    
    const cursor = doc.getCursor();
    let insertIndex = 0;
    let insertAfterCurrentParagraph = false;

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
          insertAfterCurrentParagraph = true;
        } else if (offset === fullText.length) {
          insertAfterCurrentParagraph = true;
        }
      } else {
        insertAfterCurrentParagraph = true;
      }
    }

    if (insertAfterCurrentParagraph) {
      insertIndex++;
    }

    let currentIndex = insertIndex;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const trimmed = line.trim();
      
      if (trimmed === '') {
        body.insertParagraph(currentIndex++, '');
        continue;
      }

      if (trimmed.startsWith('# ')) {
        const p = body.insertParagraph(currentIndex++, trimmed.substring(2));
        p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      } 
      else if (trimmed.startsWith('## ')) {
        const p = body.insertParagraph(currentIndex++, trimmed.substring(3));
        p.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      } 
      else if (trimmed.startsWith('### ')) {
        const p = body.insertParagraph(currentIndex++, trimmed.substring(4));
        p.setHeading(DocumentApp.ParagraphHeading.HEADING3);
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
