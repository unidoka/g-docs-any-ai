# Unidoka's GDocs any AI

GDocs any AI is a Google Docs assistant that integrates any OpenAI-compatible large language model directly into your document editor. You can generate content, rewrite highlighted text, or insert new text exactly at your cursor position. It supports vision models (images) and can optionally use your entire document as context for more accurate responses.

---

## Overview

This script adds a menu item to Google Docs. When you open the sidebar, you can ask an AI to write, expand, or rephrase text. If you highlight a piece of text, the AI will replace it with the new content. If you do not highlight anything, the response will be inserted at your current cursor location. The script also supports attaching an image (for models that can understand images) and optionally sending your full document as context (off by default to save tokens).

---

## Getting Started

Before you begin, you will need:

1. A Google account with access to Google Docs.
2. An API key from an AI provider that supports the OpenAI-compatible `/chat/completions` endpoint. This works with:
   - Google Gemini (via Google AI Studio)
   - Yandex Cloud (foundation models)
   - OpenRouter
   - OpenAI
   - Or any other provider offering a compatible endpoint.

---

## Installation (Step-by-Step)

### Step 1: Open the Apps Script Editor

1. Open the Google Doc where you want to use this tool.
2. In the top menu, go to **Extensions** > **Apps Script**.
3. This will open a new browser tab with the Apps Script editor.

### Step 2: Paste the Script

1. In the Apps Script editor, you will see a file named `Code.gs` (or similar).
2. Delete any default code in that file.
3. Copy the entire script provided in the `AppsScript.js` source code and paste it into the editor.
4. Click the **Save** icon (or press Ctrl+S / Cmd+S). Give the project a name, such as `GDocs AI`.

### Step 3: Set Up Environment Variables (Script Properties)

The script reads your API credentials from Script Properties, not from the code itself. This is more secure and easier to update.

1. In the Apps Script editor, click on the **Project Settings** tab (the gear icon) on the left sidebar.
2. Under the **Script Properties** section, click **Edit script properties** (or the "Add" button).
3. Add the following three properties. Make sure to use exactly these names:

| Property Name   | Value Example                                             | Description                                                                                                                     |
| :-------------- | :-------------------------------------------------------- | :------------------------------------------------------------------------------------------------------------------------------ |
| `AI_API_KEY`    | `your-api-key-here`                                       | The API key you obtained from your AI provider.                                                                                 |
| `AI_BASE_URL`   | `https://generativelanguage.googleapis.com/v1beta/openai` | The root URL of the AI provider's OpenAI-compatible endpoint. Do not add `/chat/completions`; the script adds it automatically. |
| `AI_MODEL_NAME` | `gemini-2.0-flash`                                        | The exact model identifier (e.g., `gemini-2.0-flash`, `gpt-4o`, or a Yandex URI like `gpt://folder/deepseek-v4-flash/latest`).  |

#### Common Configuration Examples

- **Google Gemini (default)**:
  - `AI_BASE_URL`: `https://generativelanguage.googleapis.com/v1beta/openai`
  - `AI_MODEL_NAME`: `gemini-2.0-flash`

- **Yandex Cloud**:
  - `AI_BASE_URL`: `https://ai.api.cloud.yandex.net/v1`
  - `AI_MODEL_NAME`: `gpt://b1gd47capubi1hd3o34p/deepseek-v4-flash/latest` *(replace with your actual folder ID)*

- **OpenRouter**:
  - `AI_BASE_URL`: `https://openrouter.ai/api/v1`
  - `AI_MODEL_NAME`: `google/gemini-2.0-flash-exp:free`

- **OpenAI**:
  - `AI_BASE_URL`: `https://api.openai.com/v1`
  - `AI_MODEL_NAME`: `gpt-4o`

4. After entering the values, click **Save**.

### Step 4: Save and Refresh Google Docs

1. Go back to your Google Docs tab.
2. Refresh the page (press F5 or Cmd+R).
3. Wait a few seconds for the menu to load. You will now see a new menu called **GDocs any AI** at the top of your document.

> **Note:** If you do not see the menu, open the Apps Script editor again, click **Run** > **onOpen**, and then grant the necessary permissions when prompted. Then refresh your Google Doc.

---

## How to Use the Assistant

1. From the **GDocs any AI** menu, select **Open Chat Sidebar**.
2. A sidebar will open on the right side of your document.
3. Type your instruction into the chat input field and press Enter.
4. The AI will generate a response and display it in the sidebar.
5. If you highlighted some text before sending your message, the script will show a small preview of the selected text at the top of the input area. The AI knows to replace that entire selection with its answer.
6. Click the **Insert into Document** button that appears below the AI's response. The script will delete your selected text (if any) and insert the new text directly at the cursor location, applying the document's standard styles (Times New Roman, size 14, 1.5 line spacing, 1.25 cm indent).

### Additional Controls in the Sidebar

- **Generate with images**: Toggle this on if you attach an image. Note that only vision-capable models (like Gemini 1.5 Pro) support this feature.
- **Include full document context**: Toggle this on to send the entire content of your current Google Doc as a Markdown context to the AI. This consumes significantly more tokens, so it is turned off by default.
- **Attach image**: Click the plus/paperclip icon next to the input field to upload an image. It will be sent along with your text prompt.
- **API Settings**: Use the gear menu to update your API key, model, or base URL directly from the sidebar without editing Script Properties. Changes saved here are stored permanently.

---

## Troubleshooting & Tips

- **Error: "Unexpected token 'r', 'request-id'..."**: This happens when the Base URL is incorrect or the endpoint does not return valid JSON. Make sure your `AI_BASE_URL` is the root of the OpenAI-compatible endpoint and that it automatically appends `/chat/completions` correctly (e.g., do not include `chat/completions` in the Base URL itself).
- **Text is inserting at the wrong place**: The script always inserts at the cursor position that was active at the moment you clicked "Send". If you move your cursor after generating a response but before clicking "Insert", it will still use the original location, ensuring you do not replace the wrong text.
- **Selected text is not being replaced**: Make sure you highlight text before typing your prompt. The script captures the selection at the time of sending, not when inserting. If you are unsure, look at the "Selected text" indicator in the sidebar to confirm it captured your selection.
- **Image responses**: If the AI does not generate images, your model may not support vision, or the API provider may not handle `image_url` content blocks. Check your provider's documentation for vision compatibility.
- **Token usage**: Leaving "Include full document context" enabled for long documents will cost you significantly more tokens. Turn it off unless you need the AI to understand the broader structure of the file.

---

## Updating the Script

To update the script to a newer version, simply open the Apps Script editor again, replace the existing code with the new version, save it, and refresh your Google Doc. Your Script Properties are stored separately and will be preserved.

---

## License and Disclaimer

This tool is provided as a helper utility for Google Workspace. It uses external AI APIs, and you are responsible for the usage costs and data privacy policies of your chosen provider. Always review the generated content before inserting it into important documents.