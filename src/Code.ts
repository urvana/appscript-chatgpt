/**
 * Custom function to call ChatGPT API
 *
 * @param {string} prompt The prompt to send to ChatGPT. Defaults to GT3.5 Turbo model.
 * @param {string} [model='gpt-3.5-turbo'] The model to use (e.g., 'gpt-3.5-turbo', 'gpt-4')
 * @param {number} [maxTokens=150] The maximum number of tokens to return
 * @return {string} The response from ChatGPT
 * @customfunction
 */
function CHATGPT(
  prompt: string,
  model: string = "gpt-3.5-turbo",
  maxTokens: number = 150,
): string {
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

  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
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
function CHATGPT3(prompt: string, maxTokens: number = 150): string {
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
function CHATGPT4(prompt: string, maxTokens: number = 150): string {
  return CHATGPT(prompt, "gpt-4", maxTokens);
}

/**
 * Custom function to set the OpenAI API key
 *
 * @param {string} apiKey The OpenAI API key to save
 * @return {string} Confirmation message
 * @customfunction
 */
function CHATGPTKEY(apiKey: string): string {
  PropertiesService.getUserProperties().setProperty("OPENAI_API_KEY", apiKey);
  return "✅ Remove this cell";
}
