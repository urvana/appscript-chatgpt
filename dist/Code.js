/**********************************************
 * @author Patricio López Juri
 * @license MIT
 * @version 1.1.0
 * @see {@link https://github.com/urvana/appscript-chatgpt}
 */
/** You can change this. */
const SYSTEM_PROMPT = `
  You are a helpful assistant integrated within a Google Sheets application.
  Your task is to provide accurate, concise, and user-friendly responses to user prompts.
  Explanation is not needed, just provide the best answer you can.
`;
/** Prefer short answers. ChatGPT default is 4096 */
const DEFAULT_MAX_TOKENS = 150;
/** Prefer deterministic and less creative answers. ChatGPT default is 0.7 */
const DEFAULT_TEMPERATURE = 0.3;
/** Value for empty results */
const EMPTY = "EMPTY";
/** Optional: hardcode your API key here. */
const OPENAI_API_KEY = "";
/** Private user properties storage keys. This is not the API Key itself. */
const PROPERTY_KEY_OPENAPI = "OPENAI_API_KEY";
const MIME_JSON = "application/json";
function REQUEST(
  apiKey,
  promptSystem,
  prompt,
  model = "gpt-3.5-turbo",
  maxTokens = DEFAULT_MAX_TOKENS,
  temperature = DEFAULT_TEMPERATURE,
) {
  // Prepare user prompt
  const cleaned = STRING_CLEAN(prompt);
  if (cleaned === "") {
    return EMPTY;
  }
  // Prepare system prompt
  const cleanedSystem = STRING_CLEAN(promptSystem);
  // Compose messages.
  const messages = [];
  if (cleanedSystem !== "") {
    messages.push({ role: "system", content: cleanedSystem });
  }
  messages.push({ role: "user", content: cleaned });
  const payload = {
    model: model,
    max_tokens: maxTokens,
    messages: messages,
    temperature: temperature,
  };
  const options = {
    method: "post",
    contentType: MIME_JSON,
    headers: {
      Accept: MIME_JSON,
      Authorization: `Bearer ${apiKey}`,
    },
    payload: JSON.stringify(payload),
  };
  const response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/chat/completions",
    options,
  );
  const json = response.getContentText();
  const data = JSON.parse(json);
  const choice = data["choices"][0];
  if (choice) {
    const content = choice["message"]["content"];
    return (content || "").trim() || EMPTY;
  }
  return EMPTY;
}
/**
 * Custom function to call ChatGPT API.
 * Example: =CHATGPT("What is the average height of " & A1 & "?")
 *
 * @param {string|Array<Array<string>>} prompt The prompt to send to ChatGPT.
 * @param {string} model [OPTIONAL] The model to use (e.g., "gpt-3.5-turbo", "gpt-4"). Default is "gpt-3.5-turbo" which is the most cost-effective.
 * @param {number} maxTokens [OPTIONAL] The maximum number of tokens to return. Default is 150, which is a short response. ChatGPT web default is 4096.
 * @param {number} temperature [OPTIONAL] The randomness of the response. Lower values are more deterministic. Default is 0.3 but ChatGPT web default is 0.7.
 * @return {string|Array<Array<string>>} The response from ChatGPT.
 * @customfunction
 */
function CHATGPT(
  prompt,
  model = "gpt-3.5-turbo",
  maxTokens = DEFAULT_MAX_TOKENS,
  temperature = DEFAULT_TEMPERATURE,
) {
  const properties = PropertiesService.getUserProperties();
  const apiKey = properties.getProperty(PROPERTY_KEY_OPENAPI) || OPENAI_API_KEY;
  if (!apiKey) {
    throw new Error(
      'Use =CHATGPTKEY("YOUR_API_KEY") first. Get it from https://platform.openai.com/api-keys',
    );
  }
  if (Array.isArray(prompt)) {
    return prompt.map((row) => {
      return row.map((cell) => {
        return REQUEST(
          apiKey,
          SYSTEM_PROMPT,
          cell,
          model,
          maxTokens,
          temperature,
        );
      });
    });
  }
  return REQUEST(apiKey, SYSTEM_PROMPT, prompt, model, maxTokens, temperature);
}
/**
 * Custom function to call ChatGPT-3 API. This is the default and most cost-effective model.
 * Example: =CHATGPT3("Summarize the plot of the movie: " & A1)
 *
 * @param {string|Array<Array<string>>} prompt The prompt to send to ChatGPT
 * @param {number} maxTokens [OPTIONAL] The maximum number of tokens to return. Default is 150, which is a short response. ChatGPT web default is 4096.
 * @param {number} temperature [OPTIONAL] The randomness of the response. Lower values are more deterministic. Default is 0.3 but ChatGPT web default is 0.7.
 * @return {string|Array<Array<string>>} The response from ChatGPT
 * @customfunction
 */
function CHATGPT3(
  prompt,
  maxTokens = DEFAULT_MAX_TOKENS,
  temperature = DEFAULT_TEMPERATURE,
) {
  return CHATGPT(prompt, "gpt-3.5-turbo", maxTokens, temperature);
}
/**
 * Custom function to call ChatGPT-4 API. This is the latest and most powerful model.
 * Example: =CHATGPT4("Categorize the following text into 'positive' or 'negative': " & A1)
 *
 * @param {string|Array<Array<string>>} prompt The prompt to send to ChatGPT
 * @param {number} maxTokens [OPTIONAL] The maximum number of tokens to return. Default is 150, which is a short response. ChatGPT web default is 4096.
 * @param {number} temperature [OPTIONAL] The randomness of the response. Lower values are more deterministic. Default is 0.3 but ChatGPT web default is 0.7.
 * @return {string|Array<Array<string>>} The response from ChatGPT
 * @customfunction
 */
function CHATGPT4(
  prompt,
  maxTokens = DEFAULT_MAX_TOKENS,
  temperature = DEFAULT_TEMPERATURE,
) {
  return CHATGPT(prompt, "gpt-4", maxTokens, temperature);
}
/**
 * Custom function to set the OpenAI API key. Get yours at: https://platform.openai.com/api-keys
 *
 * @param {string} apiKey The OpenAI API key to save
 * @return {string} Confirmation message
 * @customfunction
 */
function CHATGPTKEY(apiKey) {
  const properties = PropertiesService.getUserProperties();
  properties.setProperty("OPENAI_API_KEY", apiKey);
  return "✅ Remove this cell";
}
function STRING_CLEAN(value) {
  if (value === undefined || value === null) {
    return "";
  }
  if (typeof value === "number") {
    return value.toString();
  }
  return String(value).trim();
}
