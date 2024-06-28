# =CHATGPT() formula for Google Spreadsheet

This is a simple formula that allows you to use OpenAI's ChatGPT API in Google Sheets. It's a simple way to generate text using GPT-3 directly in your Google Sheets.

![demo](./assets/demo1.png)

## Installation

Here's a quick demo of how to use the formula:

https://github.com/urvana/appscript-chatgpt/assets/7570744/9293848a-803f-4d8d-8c0f-f98276f48d19

1. Open your Google Sheet.
2. Go to `Extensions` -> `Apps Script`.
3. Paste the code from [`Code.js`](https://raw.githubusercontent.com/urvana/appscript-chatgpt/main/dist/Code.js) into the script editor.
4. Save the script.
5. Use the formula `=CHATGPTKEY("YOUR_API_KEY")` in any cell.
    * Get your API key from OpenAI's website: https://platform.openai.com/api-keys
    * You can delete this cell to remove the API key from the sheet, it's saved in your personal script properties.
6. Use the formula `=CHATGPT("Hello, how are you?")` in any cell to generate text.

To compose a complex prompts concatenating multiple cells, you can use the `&` operator. For example: 

```
=CHATGPT("What is the difference between: " A1 & " and " & A2)
```

### Formulas

* `=CHATGPTKEY("YOUR_API_KEY")`: Set your API key.
* `=CHATGPT("prompt")`: Generate text based on the prompt, defaulting to **GPT-3.5-turbo**.
    * Has two optional parameters: `model=gpt-3.5-turbo` and `max_tokens=150`.
* `=CHATGPT4("prompt")`: Generate text based on the prompt using **GPT-4o**.
    * Has one optional parameter: `max_tokens=150`.
* `=CHATGPT3("prompt")`: Generate text based on the prompt using **GPT-3.5-turbo**.
    * Has one optional parameter: `max_tokens=150`.

### Code

Copy and paste the following code into the script editor:

```javascript
"use strict";
/**
 * Custom function to call ChatGPT API
 *
 * @param {string} prompt The prompt to send to ChatGPT. Defaults to GT3.5 Turbo model.
 * @param {string} [model='gpt-3.5-turbo'] The model to use (e.g., 'gpt-3.5-turbo', 'gpt-4')
 * @param {number} [maxTokens=150] The maximum number of tokens to return
 * @return {string} The response from ChatGPT
 * @customfunction
 */
function CHATGPT(prompt, model = "gpt-3.5-turbo", maxTokens = 150) {
  const apiKey =
    PropertiesService.getUserProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) {
    throw new Error(
      'API key not set. Please set the API key using the "API Key" menu.',
    );
  }
  const url = "https://api.openai.com/v1/chat/completions";
  const payload = {
    model: model,
    messages: [{ role: "user", content: prompt }],
    max_tokens: maxTokens,
  };
  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + apiKey,
    },
    payload: JSON.stringify(payload),
  };
  const response = UrlFetchApp.fetch(url, options);
  const json = response.getContentText();
  const data = JSON.parse(json);
  return data.choices[0].message.content.trim();
}
/**
 * Custom function to call ChatGPT-3 API
 *
 * @param {string} prompt The prompt to send to ChatGPT
 * @param {number} [maxTokens=150] The maximum number of tokens to return
 * @return {string} The response from ChatGPT
 * @customfunction
 */
function CHATGPT3(prompt, maxTokens = 150) {
  return CHATGPT(prompt, "gpt-3.5-turbo", maxTokens);
}
/**
 * Custom function to call ChatGPT-4 API
 *
 * @param {string} prompt The prompt to send to ChatGPT
 * @param {number} [maxTokens=150] The maximum number of tokens to return
 * @return {string} The response from ChatGPT
 * @customfunction
 */
function CHATGPT4(prompt, maxTokens = 150) {
  return CHATGPT(prompt, "gpt-4", maxTokens);
}
/**
 * Custom function to set the OpenAI API key
 *
 * @param {string} apiKey The OpenAI API key to save
 * @return {string} Confirmation message
 * @customfunction
 */
function CHATGPTKEY(apiKey) {
  PropertiesService.getUserProperties().setProperty("OPENAI_API_KEY", apiKey);
  return "✅ Remove this cell";
}
```
