/**********************************************
 * @author Patricio López Juri
 * @license MIT
 * @version 1.1.0
 * @see https://github.com/urvana/appscript-chatgpt
 */

import type {
  ChatCompletion,
  ChatCompletionCreateParamsNonStreaming,
} from "openai/resources/chat/completions";

/** You can change this. */
const SYSTEM_PROMPT = `
  You are a helpful assistant integrated within a Google Sheets application.
  Your task is to provide accurate, concise, and user-friendly responses to user prompts.
  Explanation is not needed, just provide the best answer you can.
`;

/** Value for empty results */
const EMPTY = "EMPTY" as const;
/** Optional: hardcode your API key here. */
const OPENAI_API_KEY = "";
/** Private user properties storage keys. This is not the API Key itself. */
const PROPERTY_KEY_OPENAPI = "OPENAI_API_KEY" as const;
const MIME_JSON = "application/json" as const;

type SpreadsheetInput<T> = T | Array<Array<T>>;

function REQUEST(
  apiKey: string,
  promptSystem: string,
  prompt: string,
  model = "gpt-3.5-turbo",
  maxTokens = 150,
) {
  // Prepare user prompt
  const cleaned = STRING_CLEAN(prompt);
  if (cleaned === "") {
    return EMPTY;
  }
  // Prepare system prompt
  const cleanedSystem = STRING_CLEAN(promptSystem);

  // Compose messages.
  const messages: ChatCompletionCreateParamsNonStreaming["messages"] = [];
  if (cleanedSystem !== "") {
    messages.push({ role: "system", content: cleanedSystem });
  }
  messages.push({ role: "user", content: cleaned });

  const payload: ChatCompletionCreateParamsNonStreaming = {
    model: model,
    max_tokens: maxTokens,
    messages: messages,
  };

  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
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
  const data = JSON.parse(json) as ChatCompletion;

  const choice = data["choices"][0];
  if (choice) {
    const content = choice["message"]["content"];
    return (content || "").trim() || EMPTY;
  }
  return EMPTY;
}

/**
 * Custom function to call ChatGPT API
 *
 * @param {string|Array<Array<string>>} prompt The prompt to send to ChatGPT. Defaults to GT3.5 Turbo model.
 * @param {string} [model='gpt-3.5-turbo'] The model to use (e.g., 'gpt-3.5-turbo', 'gpt-4')
 * @param {number} [maxTokens=150] The maximum number of tokens to return
 * @return {string|Array<Array<string>>} The response from ChatGPT
 * @customfunction
 */
function CHATGPT(
  prompt: SpreadsheetInput<string>,
  model = "gpt-3.5-turbo",
  maxTokens = 150,
): SpreadsheetInput<string> {
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
        return REQUEST(apiKey, SYSTEM_PROMPT, cell, model, maxTokens);
      });
    });
  }
  return REQUEST(apiKey, SYSTEM_PROMPT, prompt, model, maxTokens);
}

/**
 * Custom function to call ChatGPT-3 API
 *
 * @param {string|Array<Array<string>>} prompt The prompt to send to ChatGPT
 * @param {number} [maxTokens=150] The maximum number of tokens to return
 * @return {string|Array<Array<string>>} The response from ChatGPT
 * @customfunction
 */
function CHATGPT3(
  prompt: SpreadsheetInput<string>,
  maxTokens = 150,
): SpreadsheetInput<string> {
  return CHATGPT(prompt, "gpt-3.5-turbo", maxTokens);
}

/**
 * Custom function to call ChatGPT-4 API
 *
 * @param {string|Array<Array<string>>} prompt The prompt to send to ChatGPT
 * @param {number} [maxTokens=150] The maximum number of tokens to return
 * @return {string|Array<Array<string>>} The response from ChatGPT
 * @customfunction
 */
function CHATGPT4(
  prompt: SpreadsheetInput<string>,
  maxTokens = 150,
): SpreadsheetInput<string> {
  return CHATGPT(prompt, "gpt-4", maxTokens);
}

/**
 * Custom function to set the OpenAI API key. Get yours at: https://platform.openai.com/api-keys
 *
 * @param {string} apiKey The OpenAI API key to save
 * @return {string} Confirmation message
 * @customfunction
 */
function CHATGPTKEY(apiKey: string): string {
  const properties = PropertiesService.getUserProperties();
  properties.setProperty("OPENAI_API_KEY", apiKey);
  return "✅ Remove this cell";
}

function STRING_CLEAN(value: string | number): string {
  if (value === undefined || value === null) {
    return "";
  }
  if (typeof value === "number") {
    return value.toString();
  }
  return String(value).trim();
}
