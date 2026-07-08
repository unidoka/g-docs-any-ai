/* 
  Set up 3 env variables:
  1) Extensions -> AppScripts
  2) Project Settings tab (with the gear icon)
  3) Script Properties -> Edit script properties
  4) Add next properties
    - AI_API_KEY (e.g. YOUR_API_KEY)
    - AI_BASE_URL (e.g. https://ai.api.cloud.yandex.net/v1)
    - AI_MODEL_NAME (e.g. gpt://b1gd47capubi1hd3o34p/deepseek-v4-flash/latest)
*/

/* 
  Google Apps Script - AI Assistant 
  Universal Provider Support (OpenAI-compatible format)
  Fully English UI.
  Replaces selected text.
  Applies document styles.
*/

const DEFAULT_MODEL = "gemini-2.0-flash";
const DEFAULT_BASE_URL = "https://generativelanguage.googleapis.com/v1beta/openai";

function onOpen() {
  try {
    const ui = DocumentApp.getUi();

    const helperMenu = ui.createMenu('AI Assistant')
      .addItem('Open Chat Sidebar', 'showSidebar')
      .addSeparator()
      .addSubMenu(
        ui.createMenu('Document Tools')
          .addItem('Apply GOST Settings', 'applyGOSTSettings')
          .addItem('Create Title Page', 'createTitlePage')
          .addItem('Create Table of Contents', 'createTableOfContents')
      );
    helperMenu.addToUi();
  } catch (e) {
    Logger.log('onOpen() called from wrong context: ' + e.toString());
  }
}

function saveSettings(apiKey, modelName, baseUrl) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (apiKey && apiKey.trim()) props.setProperty('AI_API_KEY', apiKey.trim());
    if (modelName && modelName.trim()) props.setProperty('AI_MODEL_NAME', modelName.trim());
    if (baseUrl && baseUrl.trim()) props.setProperty('AI_BASE_URL', baseUrl.trim());
    return { success: true, message: 'Settings saved successfully!' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getSettings() {
  try {
    const props = PropertiesService.getScriptProperties();
    return {
      success: true,
      apiKey: props.getProperty('AI_API_KEY') || '',
      modelName: props.getProperty('AI_MODEL_NAME') || DEFAULT_MODEL,
      baseUrl: props.getProperty('AI_BASE_URL') || DEFAULT_BASE_URL
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function testConnection() {
  try {
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('AI_API_KEY');
    const modelName = props.getProperty('AI_MODEL_NAME') || DEFAULT_MODEL;
    const baseUrl = props.getProperty('AI_BASE_URL') || DEFAULT_BASE_URL;

    if (!apiKey) return { success: false, error: 'API Key missing. Please save your settings.' };

    // Standard OpenAI-compatible test payload
    const payload = {
      model: modelName,
      messages: [
        { role: "system", content: "You are a helpful assistant." },
        { role: "user", content: "Say 'OK' if you are alive." }
      ],
      temperature: 0.1,
      max_tokens: 10
    };

    const url = baseUrl + '/chat/completions';
    const authType = baseUrl.includes('yandex') ? 'Api-Key' : 'Bearer';

    const response = UrlFetchApp.fetch(url, {
      method: "post",
      headers: {
        'Authorization': authType + ' ' + apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const json = JSON.parse(response.getContentText());
    if (json.error) return { success: false, error: json.error.message };
    if (json.choices && json.choices[0]) {
      return { success: true, message: 'Connection successful! Model: ' + modelName };
    }
    return { success: false, error: 'Unexpected API response structure.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/* ---------- GOST & DOCUMENT UTILS ---------- */
function applyGOSTSettings() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    body.setMarginLeft(30 * 2.83465);
    body.setMarginRight(10 * 2.83465);
    body.setMarginTop(20 * 2.83465);
    body.setMarginBottom(20 * 2.83465);

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

    DocumentApp.getUi().alert('GOST settings applied!');
  } catch (e) {
    DocumentApp.getUi().alert('Error applying GOST: ' + e.toString());
  }
}

function createTitlePage() {
  try {
    const ui = DocumentApp.getUi();
    const response = ui.prompt(
      'Create Title Page',
      'Enter data in format (line by line):\n\n' +
      'Organization Name\nWork Title\nWork Type\nAuthor\nSupervisor\nCity\nYear',
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) return;

    const lines = response.getResponseText().split('\n').map(l => l.trim()).filter(l => l);
    if (lines.length < 7) { ui.alert('Minimum 7 lines required.'); return; }

    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    applyGOSTSettings();

    const orgPara = body.insertParagraph(0, lines[0].toUpperCase());
    orgPara.setFontFamily('Times New Roman'); orgPara.setFontSize(14);
    orgPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER); orgPara.setSpacingAfter(6);

    const titlePara = body.insertParagraph(1, lines[1].toUpperCase());
    titlePara.setFontFamily('Times New Roman'); titlePara.setFontSize(16); titlePara.setBold(true);
    titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    titlePara.setSpacingBefore(60); titlePara.setSpacingAfter(12);

    const typePara = body.insertParagraph(2, lines[2]);
    typePara.setFontFamily('Times New Roman'); typePara.setFontSize(14);
    typePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER); typePara.setSpacingAfter(40);

    const authorPara = body.insertParagraph(3, 'Done by: ' + lines[3]);
    authorPara.setFontFamily('Times New Roman'); authorPara.setFontSize(14);
    authorPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT); authorPara.setSpacingAfter(6);

    const supervisorPara = body.insertParagraph(4, 'Supervisor: ' + lines[4]);
    supervisorPara.setFontFamily('Times New Roman'); supervisorPara.setFontSize(14);
    supervisorPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT); supervisorPara.setSpacingAfter(80);

    const cityYearPara = body.insertParagraph(5, lines[5] + ', ' + lines[6]);
    cityYearPara.setFontFamily('Times New Roman'); cityYearPara.setFontSize(14);
    cityYearPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER); cityYearPara.setSpacingBefore(100);

    body.insertPageBreak(6);
    ui.alert('Title page created!');
  } catch (e) {
    DocumentApp.getUi().alert('Error: ' + e.toString());
  }
}

function createTableOfContents() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    let insertIndex = 0;
    for (let i = 0; i < body.getNumChildren(); i++) {
      if (body.getChild(i).getType() === DocumentApp.ElementType.PAGE_BREAK) {
        insertIndex = i + 1; break;
      }
    }

    const tocTitle = body.insertParagraph(insertIndex, 'TABLE OF CONTENTS');
    tocTitle.setFontFamily('Times New Roman'); tocTitle.setFontSize(14); tocTitle.setBold(true);
    tocTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER); tocTitle.setSpacingAfter(12);

    const headings = [];
    for (const para of body.getParagraphs()) {
      const h = para.getHeading();
      if (h === DocumentApp.ParagraphHeading.HEADING1 ||
        h === DocumentApp.ParagraphHeading.HEADING2 ||
        h === DocumentApp.ParagraphHeading.HEADING3) {
        headings.push({
          text: para.getText().trim(),
          level: h === DocumentApp.ParagraphHeading.HEADING1 ? 1 :
            h === DocumentApp.ParagraphHeading.HEADING2 ? 2 : 3
        });
      }
    }

    let currentIndex = insertIndex + 1;
    let sectionNumber = 1, subSectionNumber = 1;

    for (const heading of headings) {
      let prefix = '', indent = 0;
      if (heading.level === 1) {
        prefix = sectionNumber + ' '; indent = 0; subSectionNumber = 1; sectionNumber++;
      } else if (heading.level === 2) {
        prefix = sectionNumber + '.' + subSectionNumber + ' ';
        indent = 1.25 * 28.3465; subSectionNumber++;
      } else if (heading.level === 3) {
        prefix = sectionNumber + '.' + subSectionNumber + '.1 ';
        indent = 2.5 * 28.3465;
      }

      const tocItem = body.insertParagraph(currentIndex++, prefix + heading.text);
      tocItem.setFontFamily('Times New Roman'); tocItem.setFontSize(14);
      tocItem.setLineSpacing(1.5); tocItem.setIndentStart(indent); tocItem.setSpacingAfter(4);
    }

    body.insertPageBreak(currentIndex);
    DocumentApp.getUi().alert('Table of Contents created! Found ' + headings.length + ' sections.');
  } catch (e) {
    DocumentApp.getUi().alert('Error: ' + e.toString());
  }
}

/* ---------- SIDEBAR UI ---------- */
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
      --shadow-lg: 0 8px 24px rgba(0, 0, 0, 0.5);
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
      font-size: 13px; /* Increased base size */
    }

    .top-bar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 6px 8px;
      background-color: var(--bg-color);
      border-bottom: 1px solid var(--border-color);
      flex-shrink: 0;
      position: relative;
    }

    .top-bar-title {
      font-size: 12px;
      font-weight: 500;
      color: var(--text-secondary);
      display: flex;
      align-items: center;
      gap: 4px;
    }

    .top-bar-title svg {
      width: 14px;
      height: 14px;
      fill: var(--accent-color);
    }

    .params-btn {
      display: flex;
      align-items: center;
      gap: 5px;
      padding: 5px 10px;
      background-color: var(--surface-color);
      border: 1px solid var(--border-color);
      border-radius: 16px;
      cursor: pointer;
      font-size: 11px;
      font-weight: 500;
      color: var(--text-color);
      transition: all 0.2s ease;
      box-shadow: var(--shadow);
      user-select: none;
      font-family: inherit;
    }

    .params-btn:hover {
      background-color: var(--elevated-color);
      border-color: var(--accent-color);
    }

    .params-btn.active {
      background-color: var(--accent-bg);
      border-color: var(--accent-color);
      color: var(--accent-color);
    }

    .params-btn svg {
      width: 14px;
      height: 14px;
      fill: currentColor;
      transition: transform 0.25s ease;
    }

    .params-btn.active .chevron {
      transform: rotate(180deg);
    }

    .params-btn .status-dot {
      width: 6px;
      height: 6px;
      border-radius: 50%;
      background-color: var(--text-secondary);
      transition: background-color 0.2s;
    }

    .params-btn .status-dot.configured {
      background-color: var(--success-color);
      box-shadow: 0 0 4px var(--success-color);
    }

    .dropdown-panel {
      position: absolute;
      top: calc(100% + 4px);
      right: 8px;
      width: calc(100% - 16px);
      max-width: 340px;
      background-color: var(--surface-color);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      box-shadow: var(--shadow-lg);
      z-index: 100;
      opacity: 0;
      transform: translateY(-8px) scale(0.98);
      pointer-events: none;
      transition: opacity 0.18s ease, transform 0.18s ease;
      overflow: hidden;
    }

    .dropdown-panel.visible {
      opacity: 1;
      transform: translateY(0) scale(1);
      pointer-events: auto;
    }

    .dropdown-header {
      padding: 10px 12px 8px;
      border-bottom: 1px solid var(--border-color);
      display: flex;
      align-items: center;
      justify-content: space-between;
    }

    .dropdown-title {
      font-size: 12px;
      font-weight: 500;
      color: var(--text-color);
      display: flex;
      align-items: center;
      gap: 6px;
    }

    .dropdown-title svg {
      width: 14px;
      height: 14px;
      fill: var(--accent-color);
    }

    .dropdown-body {
      padding: 10px 12px;
      display: flex;
      flex-direction: column;
      gap: 10px;
    }

    .setting-item {
      display: flex;
      flex-direction: column;
      gap: 6px;
    }

    .setting-label {
      font-size: 10px;
      font-weight: 500;
      color: var(--text-secondary);
      text-transform: uppercase;
      letter-spacing: 0.5px;
      display: flex;
      align-items: center;
      gap: 4px;
    }

    .setting-label svg {
      width: 12px;
      height: 12px;
      fill: currentColor;
    }

    .toggle-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 8px 10px;
      background-color: var(--bg-color);
      border-radius: 10px;
      cursor: pointer;
      transition: background-color 0.2s;
      user-select: none;
    }

    .toggle-row:hover {
      background-color: var(--elevated-color);
    }

    .toggle-info {
      display: flex;
      align-items: center;
      gap: 8px;
      font-size: 12px;
      color: var(--text-color);
    }

    .toggle-info svg {
      width: 16px;
      height: 16px;
      fill: var(--text-secondary);
    }

    .toggle-switch {
      position: relative;
      width: 32px;
      height: 18px;
      background-color: var(--elevated-color);
      border-radius: 10px;
      transition: background-color 0.2s;
      flex-shrink: 0;
    }

    .toggle-switch::after {
      content: '';
      position: absolute;
      top: 2px;
      left: 2px;
      width: 14px;
      height: 14px;
      background-color: var(--text-secondary);
      border-radius: 50%;
      transition: all 0.2s ease;
    }

    .toggle-row.active .toggle-switch {
      background-color: var(--accent-color);
    }

    .toggle-row.active .toggle-switch::after {
      left: 16px;
      background-color: #202124;
    }

    .toggle-row.active .toggle-info svg {
      fill: var(--accent-color);
    }

    .toggle-row input { display: none; }

    .dropdown-divider {
      height: 1px;
      background-color: var(--border-color);
      margin: 2px 0;
    }

    .api-settings-btn {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 8px 10px;
      background-color: var(--bg-color);
      border: 1px solid var(--border-color);
      border-radius: 10px;
      cursor: pointer;
      transition: all 0.2s;
      color: var(--text-color);
      font-family: inherit;
      font-size: 12px;
      width: 100%;
      text-align: left;
    }

    .api-settings-btn:hover {
      background-color: var(--elevated-color);
      border-color: var(--accent-color);
    }

    .api-settings-icon {
      width: 28px;
      height: 28px;
      background-color: var(--accent-bg);
      border-radius: 8px;
      display: flex;
      align-items: center;
      justify-content: center;
      flex-shrink: 0;
    }

    .api-settings-icon svg {
      width: 16px;
      height: 16px;
      fill: var(--accent-color);
    }

    .api-settings-text {
      flex: 1;
      display: flex;
      flex-direction: column;
      gap: 1px;
    }

    .api-settings-text strong {
      font-size: 12px;
      font-weight: 500;
      color: var(--text-color);
    }

    .api-settings-text span {
      font-size: 10px;
      color: var(--text-secondary);
    }

    .api-status {
      display: flex;
      align-items: center;
      gap: 4px;
      font-size: 10px;
      color: var(--text-secondary);
    }

    .api-status.configured {
      color: var(--success-color);
    }

    .api-status-dot {
      width: 6px;
      height: 6px;
      border-radius: 50%;
      background-color: var(--text-secondary);
    }

    .api-status.configured .api-status-dot {
      background-color: var(--success-color);
      box-shadow: 0 0 4px var(--success-color);
    }

    .chevron-right {
      width: 14px;
      height: 14px;
      fill: var(--text-secondary);
      flex-shrink: 0;
    }

    .modal-overlay {
      display: none;
      position: fixed;
      top: 0; left: 0; right: 0; bottom: 0;
      background-color: rgba(0, 0, 0, 0.6);
      z-index: 1000;
      align-items: center;
      justify-content: center;
      padding: 16px;
      animation: fadeIn 0.2s ease;
    }

    .modal-overlay.visible { display: flex; }

    @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }

    .modal {
      background-color: var(--surface-color);
      border-radius: 12px;
      width: 100%;
      max-width: 360px;
      box-shadow: var(--shadow-lg);
      animation: slideUp 0.25s ease;
      overflow: hidden;
    }

    @keyframes slideUp {
      from { opacity: 0; transform: translateY(20px); }
      to { opacity: 1; transform: translateY(0); }
    }

    .modal-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 14px 16px;
      border-bottom: 1px solid var(--border-color);
    }

    .modal-title {
      font-size: 14px;
      font-weight: 500;
      color: var(--text-color);
      display: flex;
      align-items: center;
      gap: 6px;
    }

    .modal-title svg { width: 18px; height: 18px; fill: var(--accent-color); }

    .modal-close {
      background: transparent;
      border: none;
      width: 28px; height: 28px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      color: var(--text-secondary);
      transition: all 0.2s;
    }

    .modal-close:hover {
      background-color: rgba(255, 255, 255, 0.08);
      color: var(--text-color);
    }

    .modal-close svg { width: 16px; height: 16px; fill: currentColor; }

    .modal-body {
      padding: 16px;
      display: flex;
      flex-direction: column;
      gap: 14px;
    }

    .form-group {
      display: flex;
      flex-direction: column;
      gap: 4px;
    }

    .form-label {
      font-size: 11px;
      font-weight: 500;
      color: var(--text-secondary);
      display: flex;
      align-items: center;
      gap: 4px;
    }

    .form-label svg { width: 12px; height: 12px; fill: currentColor; }

    .form-input {
      width: 100%;
      padding: 8px 10px;
      background-color: var(--input-bg);
      border: 1px solid var(--border-color);
      border-radius: 8px;
      color: var(--text-color);
      font-size: 12px;
      font-family: inherit;
      outline: none;
      transition: border-color 0.2s;
    }

    .form-input:focus { border-color: var(--accent-color); }
    .form-input::placeholder { color: var(--text-secondary); opacity: 0.6; }

    .form-hint {
      font-size: 10px;
      color: var(--text-secondary);
      line-height: 1.3;
      margin-top: 2px;
    }

    .password-wrapper {
      position: relative;
      display: flex;
      align-items: center;
    }

    .password-wrapper .form-input { padding-right: 32px; }

    .password-toggle {
      position: absolute;
      right: 4px;
      background: transparent;
      border: none;
      width: 26px; height: 26px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      color: var(--text-secondary);
      transition: all 0.2s;
    }

    .password-toggle:hover {
      background-color: rgba(255, 255, 255, 0.08);
      color: var(--text-color);
    }

    .password-toggle svg { width: 14px; height: 14px; fill: currentColor; }

    .modal-footer {
      display: flex;
      align-items: center;
      justify-content: flex-end;
      gap: 8px;
      padding: 12px 16px;
      border-top: 1px solid var(--border-color);
    }

    .btn-secondary {
      padding: 6px 14px;
      background-color: transparent;
      border: 1px solid var(--border-color);
      border-radius: 8px;
      color: var(--text-secondary);
      font-size: 12px;
      font-weight: 500;
      cursor: pointer;
      transition: all 0.2s;
      font-family: inherit;
    }

    .btn-secondary:hover {
      background-color: rgba(255, 255, 255, 0.06);
      color: var(--text-color);
    }

    .btn-primary {
      padding: 6px 14px;
      background-color: var(--accent-color);
      border: none;
      border-radius: 8px;
      color: #202124;
      font-size: 12px;
      font-weight: 500;
      cursor: pointer;
      transition: all 0.2s;
      font-family: inherit;
    }

    .btn-primary:hover { background-color: var(--accent-hover); }
    .btn-primary:disabled { background-color: var(--elevated-color); color: var(--text-secondary); cursor: not-allowed; }

    .btn-test {
      padding: 6px 14px;
      background-color: rgba(129, 201, 149, 0.15);
      border: 1px solid rgba(129, 201, 149, 0.3);
      border-radius: 8px;
      color: var(--success-color);
      font-size: 12px;
      font-weight: 500;
      cursor: pointer;
      transition: all 0.2s;
      font-family: inherit;
      margin-right: auto;
    }

    .btn-test:hover { background-color: rgba(129, 201, 149, 0.25); }
    .btn-test:disabled { opacity: 0.5; cursor: not-allowed; }

    .status-message {
      display: none;
      padding: 8px 10px;
      border-radius: 8px;
      font-size: 11px;
      line-height: 1.4;
      margin-top: 4px;
    }

    .status-message.visible { display: block; }

    .status-message.success {
      background-color: rgba(129, 201, 149, 0.1);
      border: 1px solid rgba(129, 201, 149, 0.2);
      color: var(--success-color);
    }

    .status-message.error {
      background-color: rgba(242, 139, 130, 0.1);
      border: 1px solid rgba(242, 139, 130, 0.2);
      color: var(--error-color);
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
      padding: 8px 12px; /* Increased padding */
      border-radius: 8px;
      max-width: 90%; /* Increased width */
      word-wrap: break-word;
      line-height: 1.5;
      font-size: 13px; /* Larger message text */
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
      margin-top: 6px;
      padding: 6px 16px; /* Made bigger */
      background-color: var(--accent-color);
      color: #202124;
      border: none;
      border-radius: 6px;
      font-size: 12px; /* Larger button text */
      font-weight: 600;
      cursor: pointer;
      transition: background-color 0.2s, transform 0.1s;
      display: inline-block;
    }

    .insert-btn:hover { background-color: var(--accent-hover); transform: scale(1.02); }
    .insert-btn:active { transform: scale(0.95); }

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

    .chat-box:focus-within { border-color: var(--accent-color); }

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

    .mode-indicator.visible { display: flex; }
    .mode-indicator svg { width: 12px; height: 12px; fill: currentColor; flex-shrink: 0; }

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

    .selected-text-indicator.visible { display: flex; }
    .selected-text-indicator svg { width: 12px; height: 12px; fill: currentColor; flex-shrink: 0; }

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

    .attachment-preview.visible { display: flex; }

    .attachment-preview img {
      width: 32px; height: 32px;
      border-radius: 4px;
      object-fit: cover;
    }

    .attachment-info { flex: 1; min-width: 0; }

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
      width: 20px; height: 20px;
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

    .attachment-remove svg { width: 14px; height: 14px; fill: currentColor; }

    .text-input {
      width: 100%;
      background: transparent;
      border: none;
      color: var(--text-color);
      font-size: 14px;
      outline: none;
      font-family: inherit;
      padding: 2px 0;
    }

    .text-input::placeholder { color: var(--text-secondary); }

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
      width: 28px; height: 28px;
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

    .icon-btn svg { width: 18px; height: 18px; fill: currentColor; }

    .btn-send {
      background-color: var(--accent-color);
      color: #202124;
      border: none;
      width: 28px; height: 28px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      transition: background-color 0.2s, transform 0.1s;
    }

    .btn-send:hover { background-color: var(--accent-hover); }
    .btn-send:active { transform: scale(0.95); }
    .btn-send:disabled { background-color: var(--elevated-color); color: var(--text-secondary); cursor: not-allowed; }
    .btn-send svg { width: 16px; height: 16px; fill: currentColor; }

    input[type="file"] { display: none; }

    .footer-note {
      text-align: center;
      font-size: 10px;
      color: var(--text-secondary);
      padding: 2px 0 0 0;
      opacity: 0.8;
    }

    ::-webkit-scrollbar { width: 6px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: var(--elevated-color); border-radius: 3px; }
    ::-webkit-scrollbar-thumb:hover { background: var(--border-color); }
  </style>
</head>
<body>

  <div class="top-bar">
    <div class="top-bar-title">
      <svg viewBox="0 0 24 24"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 17.93c-3.95-.49-7-3.85-7-7.93 0-.62.08-1.21.21-1.79L9 15v1c0 1.1.9 2 2 2v1.93zm6.9-2.54c-.26-.81-1-1.39-1.9-1.39h-1v-3c0-.55-.45-1-1-1H8v-2h2c.55 0 1-.45 1-1V7h2c1.1 0 2-.9 2-2v-.41c2.93 1.19 5 4.06 5 7.41 0 2.08-.8 3.97-2.1 5.39z"/></svg>
      <span>AI Assistant</span>
    </div>
    
    <button class="params-btn" id="paramsBtn" onclick="toggleDropdown(event)">
      <svg viewBox="0 0 24 24"><path d="M3 17v2h6v-2H3zM3 5v2h10V5H3zm10 16v-2h8v-2h-8v-2h-2v6h2zM7 9v2H3v2h4v2h2V9H7zm14 4v-2H11v2h10zm-6-4h2V7h4V5h-4V3h-2v6z"/></svg>
      <span>Settings</span>
      <span class="status-dot" id="configDot" title="API not configured"></span>
      <svg class="chevron" viewBox="0 0 24 24"><path d="M7.41 8.59L12 13.17l4.59-4.58L18 10l-6 6-6-6 1.41-1.41z"/></svg>
    </button>
    
    <!-- DROPDOWN PANEL -->
    <div class="dropdown-panel" id="dropdownPanel">
      <div class="dropdown-header">
        <div class="dropdown-title">
          <svg viewBox="0 0 24 24"><path d="M3 17v2h6v-2H3zM3 5v2h10V5H3zm10 16v-2h8v-2h-8v-2h-2v6h2zM7 9v2H3v2h4v2h2V9H7zm14 4v-2H11v2h10zm-6-4h2V7h4V5h-4V3h-2v6z"/></svg>
          Generation Settings
        </div>
      </div>
      
      <div class="dropdown-body">
        <div class="setting-item">
          <div class="setting-label">
            <svg viewBox="0 0 24 24"><path d="M21 19V5c0-1.1-.9-2-2-2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2zM8.5 13.5l2.5 3.01L14.5 12l4.5 6H5l3.5-4.5z"/></svg>
            Images
          </div>
          <label class="toggle-row" id="imageToggle">
            <input type="checkbox" id="withImages">
            <div class="toggle-info">
              <svg viewBox="0 0 24 24"><path d="M21 19V5c0-1.1-.9-2-2-2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2zM8.5 13.5l2.5 3.01L14.5 12l4.5 6H5l3.5-4.5z"/></svg>
              <span>Generate with images</span>
            </div>
            <div class="toggle-switch"></div>
          </label>
        </div>
        
        <div class="dropdown-divider"></div>
        
        <div class="setting-item">
          <button class="api-settings-btn" onclick="openSettings()">
            <div class="api-settings-icon">
              <svg viewBox="0 0 24 24"><path d="M19.14,12.94c0.04-0.3,0.06-0.61,0.06-0.94c0-0.32-0.02-0.64-0.07-0.94l2.03-1.58c0.18-0.14,0.23-0.41,0.12-0.61 l-1.92-3.32c-0.12-0.22-0.37-0.29-0.59-0.22l-2.39,0.96c-0.5-0.38-1.03-0.7-1.62-0.94L14.4,2.81c-0.04-0.24-0.24-0.41-0.48-0.41 h-3.84c-0.24,0-0.43,0.17-0.47,0.41L9.25,5.35C8.66,5.59,8.12,5.92,7.63,6.29L5.24,5.33c-0.22-0.08-0.47,0-0.59,0.22L2.74,8.87 C2.62,9.08,2.66,9.34,2.86,9.48l2.03,1.58C4.84,11.36,4.8,11.69,4.8,12s0.02,0.64,0.07,0.94l-2.03,1.58 c-0.18,0.14-0.23,0.41-0.12,0.61l1.92,3.32c0.12,0.22,0.37,0.29,0.59,0.22l2.39-0.96c0.5,0.38,1.03,0.7,1.62,0.94l0.36,2.54 c0.05,0.24,0.24,0.41,0.48,0.41h3.84c0.24,0,0.44-0.17,0.47-0.41l0.36-2.54c0.59-0.24,1.13-0.56,1.62-0.94l2.39,0.96 c0.22,0.08,0.47,0,0.59-0.22l1.92-3.32c0.12-0.22,0.07-0.47-0.12-0.61L19.14,12.94z M12,15.6c-1.98,0-3.6-1.62-3.6-3.6 s1.62-3.6,3.6-3.6s3.6,1.62,3.6,3.6S13.98,15.6,12,15.6z"/></svg>
            </div>
            <div class="api-settings-text">
              <strong>API Settings</strong>
              <span>Key, model, and endpoint</span>
            </div>
            <div class="api-status" id="apiStatusIndicator">
              <div class="api-status-dot"></div>
              <span>Not configured</span>
            </div>
            <svg class="chevron-right" viewBox="0 0 24 24"><path d="M8.59 16.59L13.17 12 8.59 7.41 10 6l6 6-6 6-1.41-1.41z"/></svg>
          </button>
        </div>
      </div>
    </div>
  </div>

  <div class="modal-overlay" id="settingsModal">
    <div class="modal">
      <div class="modal-header">
        <div class="modal-title">
          <svg viewBox="0 0 24 24"><path d="M19.14,12.94c0.04-0.3,0.06-0.61,0.06-0.94c0-0.32-0.02-0.64-0.07-0.94l2.03-1.58c0.18-0.14,0.23-0.41,0.12-0.61 l-1.92-3.32c-0.12-0.22-0.37-0.29-0.59-0.22l-2.39,0.96c-0.5-0.38-1.03-0.7-1.62-0.94L14.4,2.81c-0.04-0.24-0.24-0.41-0.48-0.41 h-3.84c-0.24,0-0.43,0.17-0.47,0.41L9.25,5.35C8.66,5.59,8.12,5.92,7.63,6.29L5.24,5.33c-0.22-0.08-0.47,0-0.59,0.22L2.74,8.87 C2.62,9.08,2.66,9.34,2.86,9.48l2.03,1.58C4.84,11.36,4.8,11.69,4.8,12s0.02,0.64,0.07,0.94l-2.03,1.58 c-0.18,0.14-0.23,0.41-0.12,0.61l1.92,3.32c0.12,0.22,0.37,0.29,0.59,0.22l2.39-0.96c0.5,0.38,1.03,0.7,1.62,0.94l0.36,2.54 c0.05,0.24,0.24,0.41,0.48,0.41h3.84c0.24,0,0.44-0.17,0.47-0.41l0.36-2.54c0.59-0.24,1.13-0.56,1.62-0.94l2.39,0.96 c0.22,0.08,0.47,0,0.59-0.22l1.92-3.32c0.12-0.22,0.07-0.47-0.12-0.61L19.14,12.94z M12,15.6c-1.98,0-3.6-1.62-3.6-3.6 s1.62-3.6,3.6-3.6s3.6,1.62,3.6,3.6S13.98,15.6,12,15.6z"/></svg>
          AI Configuration
        </div>
        <button class="modal-close" onclick="closeSettings()">
          <svg viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>
        </button>
      </div>
      
      <div class="modal-body">
        <div class="form-group">
          <label class="form-label">
            <svg viewBox="0 0 24 24"><path d="M18 8h-1V6c0-2.76-2.24-5-5-5S7 3.24 7 6v2H6c-1.1 0-2 .9-2 2v10c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V10c0-1.1-.9-2-2-2zm-6 9c-1.1 0-2-.9-2-2s.9-2 2-2 2 .9 2 2-.9 2-2 2zm3.1-9H8.9V6c0-1.71 1.39-3.1 3.1-3.1 1.71 0 3.1 1.39 3.1 3.1v2z"/></svg>
            API Key
          </label>
          <div class="password-wrapper">
            <input type="password" id="settingsApiKey" class="form-input" placeholder="Enter API Key...">
            <button class="password-toggle" onclick="togglePasswordVisibility()" title="Show/Hide">
              <svg id="eyeIcon" viewBox="0 0 24 24"><path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/></svg>
            </button>
          </div>
          <div class="form-hint">Key for the AI provider (Gemini, OpenRouter, Yandex, etc.)</div>
        </div>
        
        <div class="form-group">
          <label class="form-label">
            <svg viewBox="0 0 24 24"><path d="M21 10.12h-6.78l2.74-2.82c-2.73-2.7-7.15-2.8-9.88-.1-2.73 2.71-2.73 7.08 0 9.79s7.15 2.71 9.88 0C18.32 15.65 19 14.08 19 12.1h2c0 1.98-.88 4.55-2.64 6.29-3.51 3.48-9.21 3.48-12.72 0-3.5-3.47-3.5-9.12 0-12.6 3.51-3.47 9.14-3.49 12.65 0L21 3v7.12z"/></svg>
            Model Name
          </label>
          <input type="text" id="settingsModelName" class="form-input" placeholder="e.g. gemini-2.0-flash">
          <div class="form-hint">Model identifier (e.g. gemini-2.0-flash, gpt-4o, or gpt://folder/model/latest)</div>
        </div>
        
        <div class="form-group">
          <label class="form-label">
            <svg viewBox="0 0 24 24"><path d="M3.9 12c0-1.71 1.39-3.1 3.1-3.1h4V7H7c-2.76 0-5 2.24-5 5s2.24 5 5 5h4v-1.9H7c-1.71 0-3.1-1.39-3.1-3.1zM8 13h8v-2H8v2zm9-6h-4v1.9h4c1.71 0 3.1 1.39 3.1 3.1s-1.39 3.1-3.1 3.1h-4V17h4c2.76 0 5-2.24 5-5s-2.24-5-5-5z"/></svg>
            Base URL
          </label>
          <input type="text" id="settingsBaseUrl" class="form-input" placeholder="https://generativelanguage.googleapis.com/v1beta/openai">
          <div class="form-hint">Must end with the API path (e.g. https://api.openrouter.ai/v1)</div>
        </div>
        
        <div class="status-message" id="settingsStatus"></div>
      </div>
      
      <div class="modal-footer">
        <button class="btn-test" id="testBtn" onclick="testApiConnection()">🔌 Test</button>
        <button class="btn-secondary" onclick="closeSettings()">Cancel</button>
        <button class="btn-primary" id="saveBtn" onclick="saveSettingsForm()">💾 Save</button>
      </div>
    </div>
  </div>

  <div class="chat-container" id="chat">
    <div class="welcome-card">
      Hello! I help you write documentation. Describe what you need to add.
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

      <input type="text" id="userInput" class="text-input" placeholder="Ask AI..." onkeypress="if(event.key==='Enter') sendMessage()">
      
      <div class="actions-row">
        <div class="left-actions">
          <button class="icon-btn" title="Attach Image" onclick="uploadImage()">
            <svg viewBox="0 0 24 24"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
          </button>
          <input type="file" id="imageInput" accept="image/*" onchange="previewImage()">
        </div>
        
        <button class="btn-send" id="sendBtn" title="Send" onclick="sendMessage()">
          <svg viewBox="0 0 24 24"><path d="M4 12l1.41 1.41L11 7.83V20h2V7.83l5.58 5.59L20 12l-8-8-8 8z"/></svg>
        </button>
      </div>
      <div class="footer-note">* AI-generated content may vary. Review before inserting.</div>
    </div>
  </div>

  <script>
    let selectedImageBase64 = null;
    let selectedImageName = '';
    
    const paramsBtn = document.getElementById('paramsBtn');
    const dropdownPanel = document.getElementById('dropdownPanel');

    function toggleDropdown(e) {
      e.stopPropagation();
      const isOpen = dropdownPanel.classList.contains('visible');
      if (isOpen) {
        closeDropdown();
      } else {
        dropdownPanel.classList.add('visible');
        paramsBtn.classList.add('active');
      }
    }

    function closeDropdown() {
      dropdownPanel.classList.remove('visible');
      paramsBtn.classList.remove('active');
    }

    document.addEventListener('click', function(e) {
      if (!dropdownPanel.contains(e.target) && !paramsBtn.contains(e.target)) {
        closeDropdown();
      }
    });

    document.addEventListener('keydown', function(e) {
      if (e.key === 'Escape') {
        const modal = document.getElementById('settingsModal');
        if (modal.classList.contains('visible')) {
          closeSettings();
        } else {
          closeDropdown();
        }
      }
    });

    function updateApiStatusIndicator(isConfigured) {
      const indicator = document.getElementById('apiStatusIndicator');
      const dot = document.getElementById('configDot');
      
      if (isConfigured) {
        indicator.classList.add('configured');
        indicator.querySelector('span').innerText = 'Configured';
        dot.classList.add('configured');
        dot.title = 'API configured';
      } else {
        indicator.classList.remove('configured');
        indicator.querySelector('span').innerText = 'Not configured';
        dot.classList.remove('configured');
        dot.title = 'API not configured';
      }
    }
    google.script.run
      .withSuccessHandler((settings) => {
        if (settings.success && settings.apiKey) {
          updateApiStatusIndicator(true);
        }
      })
      .getSettings();
    
    const imageToggle = document.getElementById('imageToggle');
    const withImagesCheckbox = document.getElementById('withImages');
    const modeIndicator = document.getElementById('modeIndicator');

    imageToggle.addEventListener('click', function(e) {
      e.stopPropagation();
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
    
    function openSettings() {
      closeDropdown();
      const modal = document.getElementById('settingsModal');
      modal.classList.add('visible');
      
      google.script.run
        .withSuccessHandler((settings) => {
          if (settings.success) {
            document.getElementById('settingsApiKey').value = settings.apiKey || '';
            document.getElementById('settingsModelName').value = settings.modelName || '';
            document.getElementById('settingsBaseUrl').value = settings.baseUrl || '';
          }
        })
        .getSettings();
      
      const status = document.getElementById('settingsStatus');
      status.classList.remove('visible', 'success', 'error');
    }

    function closeSettings() {
      document.getElementById('settingsModal').classList.remove('visible');
    }

    function togglePasswordVisibility() {
      const input = document.getElementById('settingsApiKey');
      const icon = document.getElementById('eyeIcon');
      
      if (input.type === 'password') {
        input.type = 'text';
        icon.innerHTML = '<path d="M12 7c2.76 0 5 2.24 5 5 0 .65-.13 1.26-.36 1.83l2.92 2.92c1.51-1.26 2.7-2.89 3.43-4.75-1.73-4.39-6-7.5-11-7.5-1.4 0-2.74.25-3.98.7l2.16 2.16C10.74 7.13 11.35 7 12 7zM2 4.27l2.28 2.28.46.46C3.08 8.3 1.78 10.02 1 12c1.73 4.39 6 7.5 11 7.5 1.55 0 3.03-.3 4.38-.84l.42.42L19.73 22 21 20.73 3.27 3 2 4.27zM7.53 9.8l1.55 1.55c-.05.21-.08.43-.08.65 0 1.66 1.34 3 3 3 .22 0 .44-.03.65-.08l1.55 1.55c-.67.33-1.41.53-2.2.53-2.76 0-5-2.24-5-5 0-.79.2-1.53.53-2.2zm4.31-.78l3.15 3.15.02-.16c0-1.66-1.34-3-3-3l-.17.01z"/>';
      } else {
        input.type = 'password';
        icon.innerHTML = '<path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/>';
      }
    }

    function saveSettingsForm() {
      const apiKey = document.getElementById('settingsApiKey').value.trim();
      const modelName = document.getElementById('settingsModelName').value.trim();
      const baseUrl = document.getElementById('settingsBaseUrl').value.trim();
      
      const saveBtn = document.getElementById('saveBtn');
      saveBtn.disabled = true;
      saveBtn.innerText = '';
      
      google.script.run
        .withSuccessHandler((result) => {
          saveBtn.disabled = false;
          saveBtn.innerText = 'Save';
          
          const status = document.getElementById('settingsStatus');
          if (result.success) {
            status.className = 'status-message visible success';
            status.innerText = 'Settings saved successfully!';
            updateApiStatusIndicator(!!apiKey);
            setTimeout(() => closeSettings(), 1200);
          } else {
            status.className = 'status-message visible error';
            status.innerText = 'Error: ' + result.error;
          }
        })
        .withFailureHandler((err) => {
          saveBtn.disabled = false;
          saveBtn.innerText = 'Save';
          const status = document.getElementById('settingsStatus');
          status.className = 'status-message visible error';
          status.innerText = 'Error: ' + err.message;
        })
        .saveSettings(apiKey, modelName, baseUrl);
    }

    function testApiConnection() {
      const testBtn = document.getElementById('testBtn');
      testBtn.disabled = true;
      testBtn.innerText = 'Checking...';
      
      const status = document.getElementById('settingsStatus');
      
      google.script.run
        .withSuccessHandler((result) => {
          testBtn.disabled = false;
          testBtn.innerText = 'Test';
          
          if (result.success) {
            status.className = 'status-message visible success';
            status.innerText = result.message;
            updateApiStatusIndicator(true);
          } else {
            status.className = 'status-message visible error';
            status.innerText = '' + result.error;
          }
        })
        .withFailureHandler((err) => {
          testBtn.disabled = false;
          testBtn.innerText = '🔌 Test';
          status.className = 'status-message visible error';
          status.innerText = 'Error: ' + err.message;
        })
        .testConnection();
    }

    document.getElementById('settingsModal').addEventListener('click', function(e) {
      if (e.target === this) closeSettings();
    });
    
    function uploadImage() {
      document.getElementById('imageInput').click();
    }

    function previewImage() {
      const fileInput = document.getElementById('imageInput');
      const file = fileInput.files[0];
      if (!file) { clearImage(); return; }

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
        btn.innerText = 'Insert into Document';
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

      const withImages = document.getElementById('withImages').checked;
      
      let messageText = text;
      if (selectedImageBase64) {
        messageText += (text ? '\\n\\n' : '') + '[Image: ' + selectedImageName + ']';
      }
      
      if (withImages) messageText += ' [Images]';

      addMessage(messageText, true);
      input.value = '';
      
      const imageToSend = selectedImageBase64;
      const imageNameToSend = selectedImageName;
      const hasImage = !!selectedImageBase64;
      
      if (hasImage) clearImage();
      
      document.getElementById('selectedTextIndicator').classList.remove('visible');
      
      showLoading();

      const handlerSuccess = (res) => {
        hideLoading();
        if (res.success) addMessage(res.response, false);
        else addMessage("Error: " + res.error, false, false);
      };
      const handlerFailure = (err) => {
        hideLoading();
        addMessage("Error: " + err.message, false, false);
      };

      if (hasImage) {
        google.script.run
          .withSuccessHandler(handlerSuccess)
          .withFailureHandler(handlerFailure)
          .sendChatMessageWithImage(text, imageToSend, imageNameToSend, withImages);
      } else {
        google.script.run
          .withSuccessHandler(handlerSuccess)
          .withFailureHandler(handlerFailure)
          .sendChatMessage(text, withImages);
      }
    }

    function insertResponse(btn, messageText) {
      btn.disabled = true;
      btn.innerText = 'Inserting...';
      
      google.script.run
        .withSuccessHandler((res) => {
          if (res.success) {
            btn.innerText = 'Inserted! ✓';
            btn.style.background = 'var(--success-color)';
          } else {
            btn.innerText = 'Error';
            btn.style.background = 'var(--error-color)';
          }
        })
        .withFailureHandler((err) => {
          btn.innerText = 'Error';
          btn.style.background = 'var(--error-color)';
          alert("Error: " + err.message);
        })
        .insertMarkdownToDocument(messageText);
    }
  </script>
</body>
</html>
    `;

    const html = HtmlService.createHtmlOutput(htmlContent)
      .setTitle('AI Documentation Assistant')
      .setWidth(480) // Increased width
      .setHeight(800);
    DocumentApp.getUi().showSidebar(html);
  } catch (e) {
    Logger.log('showSidebar() should be launched from Google Docs menu. Error: ' + e.toString());
  }
}

/* ---------- AI BACKEND ---------- */
function getSelectedText() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const selection = doc.getSelection();
    if (!selection) return '';

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
      if (i < elements.length - 1) selectedText += '\n';
    }
    return selectedText.trim();
  } catch (e) {
    Logger.log('Error getting selected text: ' + e.toString());
    return '';
  }
}

function sendChatMessage(userMessage, withImages) {
  try {
    const currentSection = getCurrentSectionMarkdown();
    const docStyles = getDocumentStyleGuide();
    const selectedText = getSelectedText();

    let imageInstruction = '';
    if (withImages) {
      imageInstruction = '\n\nIMPORTANT: Include 1-2 relevant images in format ![description](url). ' +
        'Use direct links from Wikimedia Commons, Unsplash or other open sources.';
    }

    let selectedTextInstruction = '';
    if (selectedText && selectedText.trim()) {
      selectedTextInstruction = '\n\nSELECTED TEXT (user highlighted this text to replace/improve):\n' +
        '```markdown\n' + selectedText + '\n```\n\n' +
        'Your task is to REPLACE this selected text with improved content based on the user request. ' +
        'Return ONLY the final text that should replace it. Do not wrap it in quotes or markdown code blocks unless it is part of the content itself.';
    }

    const systemPrompt = 'You are an expert technical documentation writer.\n\n' +
      'DOCUMENT STYLE:\n```markdown\n' + docStyles + '\n```\n\n' +
      'CURRENT SECTION:\n```markdown\n' + currentSection + '\n```\n\n' +
      'TASK: Generate full text (multiple paragraphs, subheadings, lists if needed) based on the user request. ' +
      'Strictly follow the document style. Respond ONLY with text.' +
      selectedTextInstruction + imageInstruction;

    const fullPrompt = systemPrompt + "\n\nREQUEST: " + userMessage;
    const aiResponse = callAI(fullPrompt, 8192);
    return { success: true, response: aiResponse };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function sendChatMessageWithImage(userMessage, imageBase64, imageName, withImages) {
  try {
    const currentSection = getCurrentSectionMarkdown();
    const docStyles = getDocumentStyleGuide();
    const selectedText = getSelectedText();

    let imageInstruction = '';
    if (withImages) {
      imageInstruction = '\n\nIMPORTANT: Include 1-2 relevant images in format ![description](url).';
    }

    let selectedTextInstruction = '';
    if (selectedText && selectedText.trim()) {
      selectedTextInstruction = '\n\nSELECTED TEXT:\n```markdown\n' + selectedText + '\n```\n\n' +
        'Your task is to REPLACE this selected text.';
    }

    const systemPrompt = 'You are an expert technical documentation writer with vision capabilities.\n\n' +
      'DOCUMENT STYLE:\n```markdown\n' + docStyles + '\n```\n\n' +
      'CURRENT SECTION:\n```markdown\n' + currentSection + '\n```\n\n' +
      'TASK: User uploaded image "' + imageName + '" and requests: "' + userMessage + '"\n' +
      'Generate full text for the current section. Respond ONLY with text.' +
      selectedTextInstruction + imageInstruction;

    const aiResponse = callAIWithImage(systemPrompt, imageBase64, 8192);
    return { success: true, response: aiResponse };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/* Universal AI Provider call (OpenAI format) */
function callAI(prompt, maxTokens = 8192) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('AI_API_KEY');
  const modelName = props.getProperty('AI_MODEL_NAME') || DEFAULT_MODEL;
  const baseUrl = props.getProperty('AI_BASE_URL') || DEFAULT_BASE_URL;

  if (!apiKey) throw new Error("API Key not found! Please configure in Settings → API Settings.");

  const url = baseUrl + '/chat/completions';
  const authType = baseUrl.includes('yandex') ? 'Api-Key' : 'Bearer';

  const payload = {
    model: modelName,
    messages: [{ role: "user", content: prompt }],
    temperature: 0.4,
    max_tokens: maxTokens
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      'Authorization': authType + ' ' + apiKey,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(response.getContentText());
  if (json.error) throw new Error("API Error: " + json.error.message);
  if (!json.choices || !json.choices[0] || !json.choices[0].message) {
    throw new Error("Empty response from AI provider");
  }
  return json.choices[0].message.content;
}

function callAIWithImage(prompt, imageBase64, maxTokens = 8192) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('AI_API_KEY');
  const modelName = props.getProperty('AI_MODEL_NAME') || DEFAULT_MODEL;
  const baseUrl = props.getProperty('AI_BASE_URL') || DEFAULT_BASE_URL;

  if (!apiKey) throw new Error("API Key not found! Please configure in Settings → API Settings.");

  const url = baseUrl + '/chat/completions';
  const authType = baseUrl.includes('yandex') ? 'Api-Key' : 'Bearer';

  const payload = {
    model: modelName,
    messages: [{
      role: "user",
      content: [
        { type: "text", text: prompt },
        { type: "image_url", image_url: { url: "data:image/png;base64," + imageBase64 } }
      ]
    }],
    temperature: 0.4,
    max_tokens: maxTokens
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      'Authorization': authType + ' ' + apiKey,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(response.getContentText());
  if (json.error) throw new Error("API Error: " + json.error.message);
  if (!json.choices || !json.choices[0] || !json.choices[0].message) {
    throw new Error("Empty response from AI provider");
  }
  return json.choices[0].message.content;
}

/* ---------- DOCUMENT HELPERS ---------- */
function getCurrentSectionMarkdown() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const cursor = doc.getCursor();

  if (!cursor) return getDocumentAsMarkdown();

  const element = cursor.getElement();
  let currentElement = element;
  while (currentElement.getParent() &&
    currentElement.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) {
    currentElement = currentElement.getParent();
  }
  const currentIndex = body.getChildIndex(currentElement);

  let sectionStart = 0, sectionHeading = '';
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
        sectionEnd = i; break;
      }
    }
  }

  let markdown = '';
  if (sectionHeading) markdown += '# ' + sectionHeading + '\n\n';

  for (let i = sectionStart + 1; i < sectionEnd && i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    const type = child.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const text = child.asParagraph().getText().trim();
      if (text) markdown += text + '\n\n';
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      markdown += '- ' + child.asListItem().getText().trim() + '\n';
    }
  }
  return markdown || "Section is empty.";
}

function getDocumentStyleGuide() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  let styles = { headings: [], hasBullets: false, hasNumbers: false };

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    const type = child.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const p = child.asParagraph();
      const heading = p.getHeading();
      if (heading === DocumentApp.ParagraphHeading.HEADING1) styles.headings.push({ level: 1, text: p.getText() });
      else if (heading === DocumentApp.ParagraphHeading.HEADING2) styles.headings.push({ level: 2, text: p.getText() });
      else if (heading === DocumentApp.ParagraphHeading.HEADING3) styles.headings.push({ level: 3, text: p.getText() });
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      if (child.asListItem().getGlyphType() === DocumentApp.GlyphType.BULLET) styles.hasBullets = true;
      else styles.hasNumbers = true;
    }
  }

  let styleGuide = 'Document Structure:\n';
  styles.headings.slice(0, 10).forEach(h => {
    styleGuide += '#'.repeat(h.level) + ' ' + h.text + '\n';
  });
  styleGuide += '\nFormatting Rules:\n';
  styleGuide += '- Use **bold** for key terms\n- Use *italic* for emphasis\n';
  if (styles.hasBullets) styleGuide += '- Use bulleted lists (-)\n';
  if (styles.hasNumbers) styleGuide += '- Use numbered lists (1.)\n';
  return styleGuide;
}

function getDocumentAsMarkdown() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  let markdown = "";

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    const type = child.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const p = child.asParagraph();
      const text = p.getText().trim();
      if (!text) { markdown += "\n"; continue; }
      const heading = p.getHeading();
      if (heading === DocumentApp.ParagraphHeading.HEADING1) markdown += "# " + text + "\n\n";
      else if (heading === DocumentApp.ParagraphHeading.HEADING2) markdown += "## " + text + "\n\n";
      else if (heading === DocumentApp.ParagraphHeading.HEADING3) markdown += "### " + text + "\n\n";
      else markdown += text + "\n\n";
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      const item = child.asListItem();
      const text = item.getText().trim();
      if (item.getGlyphType() === DocumentApp.GlyphType.BULLET) markdown += "- " + text + "\n";
      else markdown += "1. " + text + "\n";
    } else if (type === DocumentApp.ElementType.INLINE_IMAGE) {
      markdown += "[Image]\n\n";
    }
  }
  return markdown || "Document is empty.";
}

/* ---------- INSERT & REPLACE WITH STYLES ---------- */
function insertMarkdownToDocument(markdown) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const selection = doc.getSelection();

    // Step 1: If there is a selection, delete it completely (to replace it)
    if (selection) {
      try {
        selection.delete();
      } catch (e) {
        // Fallback if selection.delete fails (e.g. complex ranges)
        // Just proceed, we'll insert at cursor
      }
    }

    // Step 2: Insert markdown at the current cursor position
    const result = insertMarkdownAtCursor(doc, markdown);
    if (result.success) return { success: true, message: 'Text inserted successfully!' };
    else return { success: false, message: result.error };
  } catch (e) {
    return { success: false, message: 'Insert error: ' + e.toString() };
  }
}

/* Apply standard document styles (Times New Roman, 14, 1.5 spacing, etc) to a paragraph */
function applyDocumentStyleToParagraph(para, heading) {
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

function insertMarkdownAtCursor(doc, markdown) {
  try {
    if (!markdown || typeof markdown !== 'string') return { success: false, error: 'Text is empty or invalid' };
    if (!doc) return { success: false, error: 'Document not found' };
    const body = doc.getBody();
    if (!body) return { success: false, error: 'Document body not found' };

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

    if (insertAfterCurrentParagraph) insertIndex++;
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
        currentIndex = insertImageFromUrl(body, currentIndex, imageMatch[2], imageMatch[1]);
        continue;
      }

      let p = null;
      let heading = DocumentApp.ParagraphHeading.NORMAL;

      if (trimmed.startsWith('# ')) {
        p = body.insertParagraph(currentIndex++, trimmed.substring(2));
        heading = DocumentApp.ParagraphHeading.HEADING1;
        p.setHeading(heading);
      } else if (trimmed.startsWith('## ')) {
        p = body.insertParagraph(currentIndex++, trimmed.substring(3));
        heading = DocumentApp.ParagraphHeading.HEADING2;
        p.setHeading(heading);
      } else if (trimmed.startsWith('### ')) {
        p = body.insertParagraph(currentIndex++, trimmed.substring(4));
        heading = DocumentApp.ParagraphHeading.HEADING3;
        p.setHeading(heading);
      } else if (trimmed.match(/^\s*[-*+]\s/)) {
        p = body.insertListItem(currentIndex++, trimmed.replace(/^\s*[-*+]\s/, ''));
        p.setGlyphType(DocumentApp.GlyphType.BULLET);
        applyInlineFormatting(p);
      } else if (trimmed.match(/^\s*\d+\.\s/)) {
        p = body.insertListItem(currentIndex++, trimmed.replace(/^\s*\d+\.\s/, ''));
        p.setGlyphType(DocumentApp.GlyphType.NUMBER);
        applyInlineFormatting(p);
      } else {
        p = body.insertParagraph(currentIndex++, trimmed);
        applyInlineFormatting(p);
      }

      // Apply document-wide styles to the inserted paragraph
      if (p) {
        applyDocumentStyleToParagraph(p, heading);
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
  return 'image/png';
}

function insertImageFromUrl(body, index, imageUrl, altText) {
  try {
    if (!imageUrl || !imageUrl.startsWith('http')) {
      body.insertParagraph(index, '[Invalid image URL]');
      return index + 1;
    }

    const response = UrlFetchApp.fetch(imageUrl, {
      muteHttpExceptions: true, followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
    });

    if (response.getResponseCode() !== 200) {
      body.insertParagraph(index, '[Failed to load image]');
      return index + 1;
    }

    const bytes = response.getContent();
    if (!bytes || bytes.length === 0) {
      body.insertParagraph(index, '[Empty image]');
      return index + 1;
    }

    const mimeType = getMimeTypeFromUrl(imageUrl);
    const imageBlob = Utilities.newBlob(bytes, mimeType, altText || 'image');
    const image = body.insertImage(index, imageBlob);

    if (altText) {
      try { image.setAltDescription(altText); image.setAltTitle(altText); } catch (e) { }
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
    } catch (e) { }

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