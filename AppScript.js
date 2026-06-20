const GEMINI_MODEL = "gemini-3.5-flash";

function onOpen() {
  try {
    const ui = DocumentApp.getUi();
    
    const gostMenu = ui.createMenu('ГОСТ')
        .addItem('Применить настройки', 'applyGOSTSettings')
        .addItem('Создать титульный лист', 'createTitlePage')
        .addItem('Создать содержание', 'createTableOfContents');
    
    ui.createMenu('Gemini Docs')
        .addItem('Открыть чат-ассистент', 'showSidebar')
        .addSeparator()
        .addSubMenu(gostMenu)
        .addToUi();
  } catch (e) {
    Logger.log('onOpen() вызвана из неправильного контекста: ' + e.toString());
  }
}

function applyGOSTSettings() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    
    const leftMargin = 30 * 2.83465;
    const rightMargin = 10 * 2.83465;
    const topMargin = 20 * 2.83465;
    const bottomMargin = 20 * 2.83465;
    
    body.setMarginLeft(leftMargin);
    body.setMarginRight(rightMargin);
    body.setMarginTop(topMargin);
    body.setMarginBottom(bottomMargin);
    
    const paragraphs = body.getParagraphs();
    for (let i = 0; i < paragraphs.length; i++) {
      const para = paragraphs[i];
      const heading = para.getHeading();
      
      if (heading === DocumentApp.ParagraphHeading.NORMAL) {
        para.setFontFamily('Times New Roman');
        para.setFontSize(14);
        para.setLineSpacing(1.5);
        para.setSpacingAfter(0);
        para.setSpacingBefore(0);
        para.setIndentFirstLine(1.25 * 28.3465);
      }
      else if (heading === DocumentApp.ParagraphHeading.HEADING1) {
        para.setFontFamily('Times New Roman');
        para.setFontSize(14);
        para.setBold(true);
        para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        para.setLineSpacing(1.5);
        para.setSpacingBefore(12);
        para.setSpacingAfter(12);
        para.setIndentFirstLine(0);
      }
      else if (heading === DocumentApp.ParagraphHeading.HEADING2) {
        para.setFontFamily('Times New Roman');
        para.setFontSize(14);
        para.setBold(true);
        para.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        para.setLineSpacing(1.5);
        para.setSpacingBefore(8);
        para.setSpacingAfter(8);
        para.setIndentFirstLine(0);
      }
      else if (heading === DocumentApp.ParagraphHeading.HEADING3) {
        para.setFontFamily('Times New Roman');
        para.setFontSize(14);
        para.setBold(true);
        para.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        para.setLineSpacing(1.5);
        para.setSpacingBefore(6);
        para.setSpacingAfter(6);
        para.setIndentFirstLine(0);
      }
    }
    
    const footer = doc.getFooter();
    footer.clear();
    const footerPara = footer.getParagraphs()[0];
    footerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    footerPara.setFontFamily('Times New Roman');
    footerPara.setFontSize(14);
    footerPara.appendPageNumber(DocumentApp.PageNumber.CURRENT);
    
    DocumentApp.getUi().alert('Настройки ГОСТ применены:\n\n' +
      '• Поля: левое 30мм, правое 10мм, верх/низ 20мм\n' +
      '• Шрифт: Times New Roman 14pt\n' +
      '• Межстрочный интервал: 1.5\n' +
      '• Абзацный отступ: 1.25 см\n' +
      '• Нумерация страниц: внизу по центру');
      
  } catch (e) {
    DocumentApp.getUi().alert('Ошибка применения ГОСТ: ' + e.toString());
  }
}

function createTitlePage() {
  try {
    const ui = DocumentApp.getUi();
    
    const response = ui.prompt(
      'Создание титульного листа',
      'Введите данные в формате:\n\n' +
      'Организация\n' +
      'Название работы\n' +
      'Тип работы (Курсовая работа / Дипломная работа / Отчёт по практике)\n' +
      'Автор\n' +
      'Руководитель\n' +
      'Город\n' +
      'Год\n\n' +
      'Пример:\n' +
      'МГУ им. М.В. Ломоносова\n' +
      'Разработка информационной системы\n' +
      'Курсовая работа\n' +
      'Иванов И.И.\n' +
      'Петров П.П.\n' +
      'Москва\n' +
      '2026',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const inputText = response.getResponseText();
    const lines = inputText.split('\n').map(line => line.trim()).filter(line => line);
    
    if (lines.length < 7) {
      ui.alert('Недостаточно данных. Нужно минимум 7 строк.');
      return;
    }
    
    const organization = lines[0];
    const title = lines[1];
    const workType = lines[2];
    const author = lines[3];
    const supervisor = lines[4];
    const city = lines[5];
    const year = lines[6];
    
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    
    applyGOSTSettings();
    
    const orgPara = body.insertParagraph(0, organization.toUpperCase());
    orgPara.setFontFamily('Times New Roman');
    orgPara.setFontSize(14);
    orgPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    orgPara.setSpacingAfter(6);
    
    const titlePara = body.insertParagraph(1, title.toUpperCase());
    titlePara.setFontFamily('Times New Roman');
    titlePara.setFontSize(16);
    titlePara.setBold(true);
    titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    titlePara.setSpacingBefore(60);
    titlePara.setSpacingAfter(12);
    
    const typePara = body.insertParagraph(2, workType);
    typePara.setFontFamily('Times New Roman');
    typePara.setFontSize(14);
    typePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    typePara.setSpacingAfter(40);
    
    const authorPara = body.insertParagraph(3, 'Выполнил: ' + author);
    authorPara.setFontFamily('Times New Roman');
    authorPara.setFontSize(14);
    authorPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    authorPara.setSpacingAfter(6);
    
    const supervisorPara = body.insertParagraph(4, 'Руководитель: ' + supervisor);
    supervisorPara.setFontFamily('Times New Roman');
    supervisorPara.setFontSize(14);
    supervisorPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    supervisorPara.setSpacingAfter(80);
    
    const cityYearPara = body.insertParagraph(5, city + ', ' + year);
    cityYearPara.setFontFamily('Times New Roman');
    cityYearPara.setFontSize(14);
    cityYearPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    cityYearPara.setSpacingBefore(100);
    
    body.insertPageBreak(6);
    
    ui.alert('Титульный лист создан!');
    
  } catch (e) {
    DocumentApp.getUi().alert('Ошибка создания титульного листа: ' + e.toString());
  }
}

function createTableOfContents() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    
    let insertIndex = 0;
    const children = body.getNumChildren();
    
    for (let i = 0; i < children; i++) {
      const child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.PAGE_BREAK) {
        insertIndex = i + 1;
        break;
      }
    }
    
    const tocTitle = body.insertParagraph(insertIndex, 'СОДЕРЖАНИЕ');
    tocTitle.setFontFamily('Times New Roman');
    tocTitle.setFontSize(14);
    tocTitle.setBold(true);
    tocTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    tocTitle.setSpacingAfter(12);
    
    const headings = [];
    const paragraphs = body.getParagraphs();
    
    for (let i = 0; i < paragraphs.length; i++) {
      const para = paragraphs[i];
      const heading = para.getHeading();
      
      if (heading === DocumentApp.ParagraphHeading.HEADING1 ||
          heading === DocumentApp.ParagraphHeading.HEADING2 ||
          heading === DocumentApp.ParagraphHeading.HEADING3) {
        
        const text = para.getText().trim();
        const level = heading === DocumentApp.ParagraphHeading.HEADING1 ? 1 :
                      heading === DocumentApp.ParagraphHeading.HEADING2 ? 2 : 3;
        
        headings.push({ text, level });
      }
    }
    
    let currentIndex = insertIndex + 1;
    let sectionNumber = 1;
    let subSectionNumber = 1;
    
    for (const heading of headings) {
      let prefix = '';
      let indent = 0;
      
      if (heading.level === 1) {
        prefix = sectionNumber + ' ';
        indent = 0;
        subSectionNumber = 1;
      } else if (heading.level === 2) {
        prefix = sectionNumber + '.' + subSectionNumber + ' ';
        indent = 1.25 * 28.3465;
        subSectionNumber++;
      } else if (heading.level === 3) {
        prefix = sectionNumber + '.' + subSectionNumber + '.1 ';
        indent = 2.5 * 28.3465;
      }
      
      if (heading.level === 1) {
        sectionNumber++;
      }
      
      const tocItem = body.insertParagraph(currentIndex++, prefix + heading.text);
      tocItem.setFontFamily('Times New Roman');
      tocItem.setFontSize(14);
      tocItem.setLineSpacing(1.5);
      tocItem.setIndentStart(indent);
      tocItem.setSpacingAfter(4);
    }
    
    body.insertPageBreak(currentIndex);
    
    DocumentApp.getUi().alert('Содержание создано!\n\nНайдено разделов: ' + headings.length);
    
  } catch (e) {
    DocumentApp.getUi().alert('Ошибка создания содержания: ' + e.toString());
  }
}

function showSidebar() {
  try {
    const htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&display=swap" rel="stylesheet">
  <style>
    :root {
      --bg-color: #202124;
      --surface-color: #2d2e30;
      --elevated-color: #3c4043;
      --text-color: #e8eaed;
      --text-secondary: #9aa0a6;
      --card-bg: #303134;
      --card-text: #bdc1c6;
      --border-color: #3c4043;
      --input-bg: #303134;
      --icon-color: #9aa0a6;
      --icon-hover: #e8eaed;
      --accent-color: #8ab4f8;
      --accent-hover: #aecbfa;
      --accent-bg: rgba(138, 180, 248, 0.12);
      --success-color: #81c995;
      --error-color: #f28b82;
      --shadow: 0 1px 2px rgba(0, 0, 0, 0.4);
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: 'Google Sans', sans-serif;
      background-color: var(--bg-color);
      color: var(--text-color);
      display: flex;
      flex-direction: column;
      height: 100vh;
      overflow: hidden;
      font-size: 12px;
    }

    .top-options {
      display: flex;
      align-items: center;
      gap: 6px;
      padding: 6px 8px;
      background-color: var(--bg-color);
      border-bottom: 1px solid var(--border-color);
      flex-shrink: 0;
    }

    .segmented-control {
      display: flex;
      background-color: var(--surface-color);
      border-radius: 14px;
      padding: 2px;
      gap: 2px;
      box-shadow: var(--shadow);
      flex: 1;
    }

    .segment {
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 3px;
      padding: 4px 8px;
      border-radius: 12px;
      cursor: pointer;
      font-size: 11px;
      font-weight: 500;
      color: var(--text-secondary);
      transition: all 0.2s ease;
      user-select: none;
      flex: 1;
      white-space: nowrap;
    }

    .segment:hover {
      background-color: rgba(138, 180, 248, 0.08);
      color: var(--text-color);
    }

    .segment.active {
      background-color: var(--accent-bg);
      color: var(--accent-color);
    }

    .segment input { display: none; }

    .segment svg {
      width: 12px;
      height: 12px;
      fill: currentColor;
      flex-shrink: 0;
    }

    .image-toggle {
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 3px;
      padding: 4px 8px;
      background-color: var(--surface-color);
      border-radius: 12px;
      cursor: pointer;
      font-size: 11px;
      font-weight: 500;
      color: var(--text-secondary);
      transition: all 0.2s ease;
      box-shadow: var(--shadow);
      user-select: none;
      white-space: nowrap;
      position: relative;
    }

    .image-toggle:hover {
      background-color: rgba(138, 180, 248, 0.08);
      color: var(--text-color);
    }

    .image-toggle.active {
      background-color: var(--accent-color);
      color: #202124;
      box-shadow: 0 2px 6px rgba(138, 180, 248, 0.3);
    }

    .image-toggle input { display: none; }

    .image-toggle svg {
      width: 12px;
      height: 12px;
      fill: currentColor;
    }

    .image-toggle::after {
      content: '';
      position: absolute;
      top: 3px;
      right: 3px;
      width: 5px;
      height: 5px;
      border-radius: 50%;
      background-color: transparent;
      transition: background-color 0.2s;
    }

    .image-toggle.active::after {
      background-color: var(--success-color);
    }

    .chat-container {
      flex: 1;
      padding: 8px;
      overflow-y: auto;
      display: flex;
      flex-direction: column;
      gap: 6px;
      min-height: 0;
    }

    .welcome-card {
      background-color: var(--card-bg);
      border-radius: 8px;
      padding: 8px 10px;
      font-size: 12px;
      line-height: 1.4;
      color: var(--card-text);
      border: 1px solid var(--border-color);
    }

    .message {
      padding: 6px 10px;
      border-radius: 8px;
      max-width: 85%;
      word-wrap: break-word;
      line-height: 1.4;
      font-size: 12px;
      animation: slideIn 0.2s ease;
    }

    @keyframes slideIn {
      from { opacity: 0; transform: translateY(8px); }
      to { opacity: 1; transform: translateY(0); }
    }

    .user {
      background-color: var(--accent-color);
      color: #202124;
      margin-left: auto;
      align-self: flex-end;
      font-weight: 500;
    }

    .ai {
      background-color: var(--card-bg);
      color: var(--card-text);
      margin-right: auto;
      align-self: flex-start;
      border: 1px solid var(--border-color);
    }

    .insert-btn {
      margin-top: 4px;
      padding: 3px 8px;
      background-color: var(--accent-color);
      color: #202124;
      border: none;
      border-radius: 4px;
      font-size: 10px;
      font-weight: 500;
      cursor: pointer;
      transition: background-color 0.2s;
    }

    .insert-btn:hover {
      background-color: var(--accent-hover);
    }

    .loading {
      text-align: center;
      color: var(--text-secondary);
      font-style: italic;
      padding: 8px;
      font-size: 11px;
    }

    .input-panel {
      padding: 6px 8px 8px 8px;
      background-color: var(--bg-color);
      flex-shrink: 0;
    }

    .chat-box {
      background-color: var(--input-bg);
      border-radius: 16px;
      padding: 6px 10px;
      display: flex;
      flex-direction: column;
      gap: 4px;
      border: 1px solid var(--border-color);
      transition: border-color 0.2s;
    }

    .chat-box:focus-within {
      border-color: var(--accent-color);
    }

    .mode-indicator {
      display: none;
      align-items: center;
      gap: 4px;
      padding: 4px 8px;
      background-color: rgba(138, 180, 248, 0.1);
      border: 1px solid rgba(138, 180, 248, 0.2);
      border-radius: 8px;
      font-size: 10px;
      color: var(--accent-color);
    }

    .mode-indicator.visible {
      display: flex;
    }

    .mode-indicator svg {
      width: 12px;
      height: 12px;
      fill: currentColor;
      flex-shrink: 0;
    }

    .selected-text-indicator {
      display: none;
      align-items: center;
      gap: 4px;
      padding: 4px 8px;
      background-color: rgba(129, 201, 149, 0.1);
      border: 1px solid rgba(129, 201, 149, 0.2);
      border-radius: 8px;
      font-size: 10px;
      color: var(--success-color);
    }

    .selected-text-indicator.visible {
      display: flex;
    }

    .selected-text-indicator svg {
      width: 12px;
      height: 12px;
      fill: currentColor;
      flex-shrink: 0;
    }

    .selected-text-content {
      flex: 1;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      font-size: 10px;
    }

    .attachment-preview {
      display: none;
      align-items: center;
      gap: 6px;
      padding: 4px 6px;
      background-color: var(--elevated-color);
      border-radius: 8px;
    }

    .attachment-preview.visible {
      display: flex;
    }

    .attachment-preview img {
      width: 32px;
      height: 32px;
      border-radius: 4px;
      object-fit: cover;
    }

    .attachment-info {
      flex: 1;
      min-width: 0;
    }

    .attachment-name {
      font-size: 11px;
      color: var(--text-color);
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .attachment-remove {
      background: transparent;
      border: none;
      color: var(--text-secondary);
      cursor: pointer;
      width: 20px;
      height: 20px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      transition: background-color 0.2s, color 0.2s;
    }

    .attachment-remove:hover {
      background-color: rgba(255, 255, 255, 0.08);
      color: var(--error-color);
    }

    .attachment-remove svg {
      width: 14px;
      height: 14px;
      fill: currentColor;
    }

    .text-input {
      width: 100%;
      background: transparent;
      border: none;
      color: var(--text-color);
      font-size: 13px;
      outline: none;
      font-family: inherit;
      padding: 2px 0;
    }

    .text-input::placeholder {
      color: var(--text-secondary);
    }

    .actions-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
    }

    .left-actions {
      display: flex;
      gap: 2px;
      align-items: center;
    }

    .icon-btn {
      background: transparent;
      border: none;
      width: 28px;
      height: 28px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      color: var(--icon-color);
      transition: background-color 0.2s, color 0.2s;
    }

    .icon-btn:hover {
      background-color: rgba(255, 255, 255, 0.08);
      color: var(--icon-hover);
    }

    .icon-btn svg {
      width: 18px;
      height: 18px;
      fill: currentColor;
    }

    .btn-send {
      background-color: var(--accent-color);
      color: #202124;
      border: none;
      width: 28px;
      height: 28px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      transition: background-color 0.2s, transform 0.1s;
    }

    .btn-send:hover {
      background-color: var(--accent-hover);
    }

    .btn-send:active {
      transform: scale(0.95);
    }

    .btn-send:disabled {
      background-color: var(--elevated-color);
      color: var(--text-secondary);
      cursor: not-allowed;
    }

    .btn-send svg {
      width: 16px;
      height: 16px;
      fill: currentColor;
    }

    input[type="file"] { display: none; }

    ::-webkit-scrollbar { width: 6px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: var(--elevated-color); border-radius: 3px; }
    ::-webkit-scrollbar-thumb:hover { background: var(--border-color); }
  </style>
</head>
<body>

  <div class="top-options">
    <div class="segmented-control">
      <label class="segment active">
        <input type="radio" name="genType" value="full" checked>
        <svg viewBox="0 0 24 24"><path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/></svg>
        <span>Full</span>
      </label>
      <label class="segment">
        <input type="radio" name="genType" value="paragraph">
        <svg viewBox="0 0 24 24"><path d="M14 17H4v2h10v-2zm6-8H4v2h16V9zM4 15h16v-2H4v2zM4 5v2h16V5H4z"/></svg>
        <span>Paragraph</span>
      </label>
      <label class="segment">
        <input type="radio" name="genType" value="sentence">
        <svg viewBox="0 0 24 24"><path d="M4 9h16v2H4V9zm0 4h10v2H4v-2z"/></svg>
        <span>Sentence</span>
      </label>
    </div>
    
    <label class="image-toggle" id="imageToggle" title="Включить генерацию изображений">
      <input type="checkbox" id="withImages">
      <svg viewBox="0 0 24 24"><path d="M21 19V5c0-1.1-.9-2-2-2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2zM8.5 13.5l2.5 3.01L14.5 12l4.5 6H5l3.5-4.5z"/></svg>
      <span>Images</span>
    </label>
  </div>

  <div class="chat-container" id="chat">
    <div class="welcome-card">
      Hi! I'll help you create documentation. Describe what you'd like to add.
    </div>
  </div>

  <div class="input-panel">
    <div class="chat-box">
      <div class="mode-indicator" id="modeIndicator">
        <svg viewBox="0 0 24 24"><path d="M21 19V5c0-1.1-.9-2-2-2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2zM8.5 13.5l2.5 3.01L14.5 12l4.5 6H5l3.5-4.5z"/></svg>
        <span>AI will include images</span>
      </div>
      
      <div class="selected-text-indicator" id="selectedTextIndicator">
        <svg viewBox="0 0 24 24"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
        <span class="selected-text-content" id="selectedTextContent"></span>
      </div>
      
      <div class="attachment-preview" id="attachmentPreview">
        <img id="attachmentThumb" src="" alt="">
        <div class="attachment-info">
          <div class="attachment-name" id="attachmentName"></div>
        </div>
        <button class="attachment-remove" onclick="clearImage()" title="Remove">
          <svg viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>
        </button>
      </div>

      <input type="text" id="userInput" class="text-input" placeholder="Ask Gemini..." onkeypress="if(event.key==='Enter') sendMessage()">
      
      <div class="actions-row">
        <div class="left-actions">
          <button class="icon-btn" title="Attach image" onclick="uploadImage()">
            <svg viewBox="0 0 24 24"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
          </button>
          <input type="file" id="imageInput" accept="image/*" onchange="previewImage()">
        </div>
        
        <button class="btn-send" id="sendBtn" title="Send" onclick="sendMessage()">
          <svg viewBox="0 0 24 24"><path d="M4 12l1.41 1.41L11 7.83V20h2V7.83l5.58 5.59L20 12l-8-8-8 8z"/></svg>
        </button>
      </div>
    </div>
  </div>

  <script>
    let selectedImageBase64 = null;
    let selectedImageName = '';

    document.querySelectorAll('.segment').forEach(segment => {
      segment.addEventListener('click', function() {
        document.querySelectorAll('.segment').forEach(s => s.classList.remove('active'));
        this.classList.add('active');
        this.querySelector('input').checked = true;
      });
    });

    const imageToggle = document.getElementById('imageToggle');
    const withImagesCheckbox = document.getElementById('withImages');
    const modeIndicator = document.getElementById('modeIndicator');

    imageToggle.addEventListener('click', function() {
      setTimeout(() => {
        if (withImagesCheckbox.checked) {
          imageToggle.classList.add('active');
          modeIndicator.classList.add('visible');
        } else {
          imageToggle.classList.remove('active');
          modeIndicator.classList.remove('visible');
        }
      }, 0);
    });

    document.getElementById('userInput').addEventListener('focus', checkSelectedText);

    function checkSelectedText() {
      google.script.run
        .withSuccessHandler((selectedText) => {
          const indicator = document.getElementById('selectedTextIndicator');
          const content = document.getElementById('selectedTextContent');
          
          if (selectedText && selectedText.trim()) {
            const preview = selectedText.substring(0, 50) + (selectedText.length > 50 ? '...' : '');
            content.innerText = preview;
            indicator.classList.add('visible');
          } else {
            indicator.classList.remove('visible');
          }
        })
        .getSelectedText();
    }

    function uploadImage() {
      document.getElementById('imageInput').click();
    }

    function previewImage() {
      const fileInput = document.getElementById('imageInput');
      const file = fileInput.files[0];
      
      if (!file) {
        clearImage();
        return;
      }

      const reader = new FileReader();
      reader.onload = function(e) {
        selectedImageBase64 = e.target.result.split(',')[1];
        selectedImageName = file.name;
        
        document.getElementById('attachmentThumb').src = e.target.result;
        document.getElementById('attachmentName').innerText = file.name;
        document.getElementById('attachmentPreview').classList.add('visible');
      };
      reader.readAsDataURL(file);
    }

    function clearImage() {
      selectedImageBase64 = null;
      selectedImageName = '';
      document.getElementById('imageInput').value = '';
      document.getElementById('attachmentPreview').classList.remove('visible');
      document.getElementById('attachmentThumb').src = '';
      document.getElementById('attachmentName').innerText = '';
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
        btn.innerText = 'Insert';
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
      div.innerText = 'Generating...';
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
      const withImages = document.getElementById('withImages').checked;
      
      let messageText = text;
      if (selectedImageBase64) {
        messageText += (text ? '\\n\\n' : '') + '[Image: ' + selectedImageName + ']';
      }
      
      const genTypeLabels = {
        'full': 'Full',
        'paragraph': 'Paragraph',
        'sentence': 'Sentence'
      };
      messageText += ' [' + genTypeLabels[genType] + ']';
      if (withImages) messageText += ' [Images]';

      addMessage(messageText, true);
      input.value = '';
      
      const imageToSend = selectedImageBase64;
      const imageNameToSend = selectedImageName;
      const hasImage = !!selectedImageBase64;
      
      if (hasImage) clearImage();
      
      document.getElementById('selectedTextIndicator').classList.remove('visible');
      
      showLoading();

      if (hasImage) {
        google.script.run
          .withSuccessHandler((res) => {
            hideLoading();
            if (res.success) {
              addMessage(res.response, false);
            } else {
              addMessage("Error: " + res.error, false, false);
            }
          })
          .withFailureHandler((err) => {
            hideLoading();
            addMessage("Error: " + err.message, false, false);
          })
          .sendChatMessageWithImage(text, imageToSend, imageNameToSend, genType, withImages);
      } else {
        google.script.run
          .withSuccessHandler((res) => {
            hideLoading();
            if (res.success) {
              addMessage(res.response, false);
            } else {
              addMessage("Error: " + res.error, false, false);
            }
          })
          .withFailureHandler((err) => {
            hideLoading();
            addMessage("Error: " + err.message, false, false);
          })
          .sendChatMessage(text, genType, withImages);
      }
    }

    function insertResponse(btn, messageText) {
      btn.disabled = true;
      btn.innerText = '⏳';
      
      google.script.run
        .withSuccessHandler((res) => {
          if (res.success) {
            btn.innerText = '✅';
            btn.style.background = 'var(--success-color)';
          } else {
            btn.innerText = '❌';
            btn.style.background = 'var(--error-color)';
          }
        })
        .withFailureHandler((err) => {
          btn.innerText = '❌';
          btn.style.background = 'var(--error-color)';
          alert("Failed: " + err.message);
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
  } catch (e) {
    Logger.log('showSidebar() should be launched from Google Docs menu. Error: ' + e.toString());
  }
}

function getSelectedText() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const selection = doc.getSelection();
    
    if (!selection) {
      return '';
    }
    
    const elements = selection.getSelectedElements();
    let selectedText = '';
    
    for (let i = 0; i < elements.length; i++) {
      const element = elements[i].getElement();
      
      if (element.getType() === DocumentApp.ElementType.TEXT) {
        const textElement = element.asText();
        const startOffset = elements[i].getStartOffset();
        const endOffset = elements[i].getEndOffsetInclusive();
        
        if (startOffset >= 0 && endOffset >= 0) {
          selectedText += textElement.getText().substring(startOffset, endOffset + 1);
        } else {
          selectedText += textElement.getText();
        }
      } else if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
        selectedText += element.asParagraph().getText();
      }
      
      if (i < elements.length - 1) {
        selectedText += '\n';
      }
    }
    
    return selectedText.trim();
  } catch (e) {
    Logger.log('Error getting selected text: ' + e.toString());
    return '';
  }
}

function sendChatMessage(userMessage, genType, withImages) {
  try {
    const currentSection = getCurrentSectionMarkdown();
    const docStyles = getDocumentStyleGuide();
    const selectedText = getSelectedText();
    
    let genTypeText = '';
    let maxTokens = 8192;
    
    if (genType === 'full') {
      genTypeText = 'full text for the current section (multiple paragraphs, subheadings, lists if needed)';
      maxTokens = 8192;
    } else if (genType === 'paragraph') {
      genTypeText = 'one paragraph (3-5 sentences)';
      maxTokens = 1024;
    } else if (genType === 'sentence') {
      genTypeText = '1-2 sentences';
      maxTokens = 256;
    }
    
    let imageInstruction = '';
    if (withImages) {
      imageInstruction = '\n\nIMPORTANT: Include 1-2 relevant images in format ![description](url). ' +
      'Use direct links from Wikimedia Commons, Unsplash or other open sources. ' +
      'Example: ![Architecture Diagram](https://upload.wikimedia.org/wikipedia/commons/thumb/.../image.png)';
    }

    let selectedTextInstruction = '';
    if (selectedText && selectedText.trim()) {
      selectedTextInstruction = '\n\nSELECTED TEXT (user highlighted this in the document):\n' +
      '```markdown\n' + selectedText + '\n```\n\n' +
      'IMPORTANT: The user has selected text in the document. Consider this text in your response. ' +
      'Analyze the user request and act accordingly:\n' +
      '- If the user asks to rewrite/improve: rewrite the selected text\n' +
      '- If the user asks to continue: continue the thought from the selected text\n' +
      '- If the user asks to expand: generate additional content based on the selected text\n' +
      '- If the request is unrelated to the selected text: still consider it as context';
    }

    const systemPrompt = 'You are an expert technical documentation writer.\n\n' +
    'DOCUMENT STYLE (copy this formatting):\n' +
    '```markdown\n' + docStyles + '\n```\n\n' +
    'CURRENT SECTION (context):\n' +
    '```markdown\n' + currentSection + '\n```\n\n' +
    'TASK:\n' +
    'Generate ' + genTypeText + ' for the current section based on the user request. ' +
    'Strictly follow the document style. Respond ONLY with text without markdown formatting (except **bold**, *italic*' + (withImages ? ' and ![images](url)' : '') + ').' +
    selectedTextInstruction +
    imageInstruction;

    const fullPrompt = systemPrompt + "\n\nREQUEST: " + userMessage;
    
    const aiResponse = callGemini(fullPrompt, maxTokens);
    
    return { success: true, response: aiResponse };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function sendChatMessageWithImage(userMessage, imageBase64, imageName, genType, withImages) {
  try {
    const currentSection = getCurrentSectionMarkdown();
    const docStyles = getDocumentStyleGuide();
    const selectedText = getSelectedText();
    
    let genTypeText = '';
    let maxTokens = 8192;
    
    if (genType === 'full') {
      genTypeText = 'full text for the current section (multiple paragraphs, subheadings, lists if needed)';
      maxTokens = 8192;
    } else if (genType === 'paragraph') {
      genTypeText = 'one paragraph (3-5 sentences)';
      maxTokens = 1024;
    } else if (genType === 'sentence') {
      genTypeText = '1 sentences';
      maxTokens = 256;
    }
    
    let imageInstruction = '';
    if (withImages) {
      imageInstruction = '\n\nIMPORTANT: Include 1-2 relevant images in format ![description](url). ' +
      'Use direct links from Wikimedia Commons, Unsplash or other open sources.';
    }

    let selectedTextInstruction = '';
    if (selectedText && selectedText.trim()) {
      selectedTextInstruction = '\n\nSELECTED TEXT (user highlighted this in the document):\n' +
      '```markdown\n' + selectedText + '\n```\n\n' +
      'IMPORTANT: The user has selected text in the document. Consider this text in your response along with the uploaded image. ' +
      'Analyze the user request and act accordingly:\n' +
      '- If the user asks to rewrite/improve: rewrite the selected text\n' +
      '- If the user asks to continue: continue the thought from the selected text\n' +
      '- If the user asks to expand: generate additional content based on the selected text and image\n' +
      '- If the request is unrelated to the selected text: still consider it as context';
    }

    const systemPrompt = 'You are an expert technical documentation writer with vision capabilities.\n\n' +
    'DOCUMENT STYLE:\n' +
    '```markdown\n' + docStyles + '\n```\n\n' +
    'CURRENT SECTION:\n' +
    '```markdown\n' + currentSection + '\n```\n\n' +
    'TASK:\n' +
    'User uploaded image "' + imageName + '" and requests: "' + userMessage + '"\n' +
    'Analyze the image and generate ' + genTypeText + ' for the current section. ' +
    'Follow the document style. Respond ONLY with text.' +
    selectedTextInstruction +
    imageInstruction;

    const aiResponse = callGeminiWithImage(systemPrompt, imageBase64, maxTokens);
    
    return { success: true, response: aiResponse };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getCurrentSectionMarkdown() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const cursor = doc.getCursor();
  
  if (!cursor) {
    return getDocumentAsMarkdown();
  }
  
  const element = cursor.getElement();
  let currentElement = element;
  
  while (currentElement.getParent() && 
         currentElement.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) {
    currentElement = currentElement.getParent();
  }
  
  const currentIndex = body.getChildIndex(currentElement);
  
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
  
  return markdown || "Section is empty.";
}

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
  
  let styleGuide = 'Document Structure:\n';
  styles.headings.slice(0, 10).forEach(h => {
    const prefix = '#'.repeat(h.level);
    styleGuide += prefix + ' ' + h.text + '\n';
  });
  
  styleGuide += '\nFormatting Rules:\n';
  styleGuide += '- Use **bold** for key terms\n';
  styleGuide += '- Use *italic* for emphasis\n';
  if (styles.hasBullets) styleGuide += '- Use bulleted lists (-)\n';
  if (styles.hasNumbers) styleGuide += '- Use numbered lists (1.)\n';
  
  return styleGuide;
}

function callGemini(prompt, maxTokens = 8192) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    throw new Error("API key not found! Add 'GEMINI_API_KEY' to Script Properties.");
  }

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + apiKey;
  
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { 
      temperature: 0.4,
      maxOutputTokens: maxTokens
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
    throw new Error("Gemini API Error: " + json.error.message);
  }
  
  if (!json.candidates || !json.candidates[0] || !json.candidates[0].content) {
    throw new Error("Empty response from Gemini");
  }
  
  return json.candidates[0].content.parts[0].text;
}

function callGeminiWithImage(prompt, imageBase64, maxTokens = 8192) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    throw new Error("API key not found!");
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
      maxOutputTokens: maxTokens
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
    throw new Error("Gemini API Error: " + json.error.message);
  }
  
  if (!json.candidates || !json.candidates[0] || !json.candidates[0].content) {
    throw new Error("Empty response from Gemini");
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
      markdown += "[Image]\n\n";
    }
  }
  
  return markdown || "Document is empty.";
}

function insertMarkdownToDocument(markdown) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const result = insertMarkdownAtCursor(doc, markdown);
    
    if (result.success) {
      return { success: true, message: 'Text inserted successfully!' };
    } else {
      return { success: false, message: result.error };
    }
  } catch (e) {
    return { success: false, message: 'Insert error: ' + e.toString() };
  }
}

function insertMarkdownAtCursor(doc, markdown) {
  try {
    if (!markdown || typeof markdown !== 'string') {
      return { success: false, error: 'Text is empty or invalid' };
    }

    if (!doc) {
      return { success: false, error: 'Document not found' };
    }

    const body = doc.getBody();
    if (!body) {
      return { success: false, error: 'Document body not found' };
    }
    
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

      const imageMatch = trimmed.match(/^!\[([^\]]*)\]\(([^)]+)\)$/);
      if (imageMatch) {
        const alt = imageMatch[1];
        const imageUrl = imageMatch[2];
        currentIndex = insertImageFromUrl(body, currentIndex, imageUrl, alt);
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
    Logger.log('Error in insertMarkdownAtCursor: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function getMimeTypeFromUrl(url) {
  const urlLower = url.toLowerCase().split('?')[0];
  
  if (urlLower.endsWith('.png')) return 'image/png';
  if (urlLower.endsWith('.jpg') || urlLower.endsWith('.jpeg')) return 'image/jpeg';
  if (urlLower.endsWith('.gif')) return 'image/gif';
  if (urlLower.endsWith('.webp')) return 'image/webp';
  if (urlLower.endsWith('.svg')) return 'image/svg+xml';
  if (urlLower.endsWith('.bmp')) return 'image/bmp';
  
  return 'image/png';
}

function insertImageFromUrl(body, index, imageUrl, altText) {
  try {
    if (!imageUrl || !imageUrl.startsWith('http')) {
      body.insertParagraph(index, '[Invalid image URL]');
      return index + 1;
    }

    const response = UrlFetchApp.fetch(imageUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      body.insertParagraph(index, '[Failed to load image: code ' + responseCode + ']');
      return index + 1;
    }

    const bytes = response.getContent();
    
    if (!bytes || bytes.length === 0) {
      body.insertParagraph(index, '[Empty image]');
      return index + 1;
    }

    const mimeType = getMimeTypeFromUrl(imageUrl);
    const imageBlob = Utilities.newBlob(bytes, mimeType, altText || 'image');
    
    if (!imageBlob || !imageBlob.getBytes() || imageBlob.getBytes().length === 0) {
      body.insertParagraph(index, '[Error creating image]');
      return index + 1;
    }

    const image = body.insertImage(index, imageBlob);
    
    if (altText) {
      try {
        image.setAltDescription(altText);
        image.setAltTitle(altText);
      } catch (e) {
        Logger.log('Failed to set alt: ' + e.toString());
      }
    }
    
    try {
      const originalWidth = image.getWidth();
      const originalHeight = image.getHeight();
      const maxWidth = 450;
      
      if (originalWidth > maxWidth) {
        const ratio = maxWidth / originalWidth;
        image.setWidth(maxWidth);
        image.setHeight(Math.round(originalHeight * ratio));
      }
    } catch (e) {
      Logger.log('Failed to resize: ' + e.toString());
    }
    
    return index + 1;
  } catch (e) {
    Logger.log('Critical error loading image: ' + e.toString());
    body.insertParagraph(index, '[Error loading image]');
    return index + 1;
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
