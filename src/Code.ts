/**********************************************
 * @author Patricio López Juri <https://www.linkedin.com/in/lopezjuri/>
 * @license MIT
 * @version 1.2.0
 * @see {@link https://github.com/urvana/appscript-chatgpt}
 */

import type {
  ChatCompletion,
  ChatCompletionCreateParamsNonStreaming,
} from "openai/resources/chat/completions";
import type { ModelsPage } from "openai/resources/models";

/** You can change this. */
const SYSTEM_PROMPT = `
  You are a helpful assistant integrated within a Google Sheets application.
  Your task is to provide accurate, concise, and user-friendly responses to user prompts.
  Explanation is not needed, just provide the best answer you can.
`;
/** Prefer short answers. ChatGPT web default is 4096 */
const DEFAULT_MAX_TOKENS = 150;
/** Prefer deterministic and less creative answers. ChatGPT web default is 0.7 */
const DEFAULT_TEMPERATURE = 0.0;
/**
 * Setup how long should responses be cached in seconds. Default is 6 hours.
 * Set to 0, undefined or null to disable caching.
 * Set to -1 to cache indefinitely.
 */
const DEFAULT_CACHE_DURATION: number | undefined | null = 21600; // 6 hours
/** Helps determinism. */
const DEFAULT_SEED = 0;

/** Value for empty results */
const EMPTY = "EMPTY" as const;
/** Optional: you can hardcode your API key here. Keep in mind it's not secure and other users can see it. */
const OPENAI_API_KEY = "";
/** Private user properties storage keys. This is not the API Key itself. */
const PROPERTY_KEY_OPENAPI = "OPENAI_API_KEY" as const;
const MIME_JSON = "application/json" as const;

type SpreadsheetInput<T> = T | Array<Array<T>>;

function REQUEST_COMPLETIONS(
  apiKey: string,
  promptSystem: string,
  prompt: string,
  model = "gpt-3.5-turbo",
  maxTokens = DEFAULT_MAX_TOKENS,
  temperature = DEFAULT_TEMPERATURE,
) {
  // Prepare user prompt
  const promptCleaned = STRING_CLEAN(prompt);
  if (promptCleaned === "") {
    return EMPTY;
  }
  // Prepare system prompt
  const promptSystemCleaned = STRING_CLEAN(promptSystem);

  // Create cache key
  const cache = GET_CACHE();
  const cacheKey = HASH_SHA1(promptCleaned, model, maxTokens, temperature);

  // Check cache
  if (DEFAULT_CACHE_DURATION) {
    const cachedResponse = cache.get(cacheKey);
    if (cachedResponse) {
      return cachedResponse;
    }
  }

  // Compose messages.
  const messages: ChatCompletionCreateParamsNonStreaming["messages"] = [];
  if (promptSystemCleaned) {
    messages.push({ role: "system", content: promptSystemCleaned });
  }
  messages.push({ role: "user", content: promptCleaned });

  /**
   * Unique user ID for the current user (rotates every 30 days).
   * https://developers.google.com/apps-script/reference/base/session?hl#getTemporaryActiveUserKey()
   */
  const user_id = Session.getTemporaryActiveUserKey();

  const payload: ChatCompletionCreateParamsNonStreaming = {
    stream: false,
    model: model,
    max_tokens: maxTokens,
    messages: messages,
    temperature: temperature,
    service_tier: "auto", // Handle rate limits automatically.
    n: 1, // Number of completions to generate.
    user: user_id, // User ID to associate with the conversation and let OpenAI handle abusers.
    seed: DEFAULT_SEED,
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
    const content = (choice["message"]["content"] || "").trim();
    if (content && DEFAULT_CACHE_DURATION === -1) {
      cache.put(cacheKey, content, Number.POSITIVE_INFINITY);
    } else if (content && DEFAULT_CACHE_DURATION) {
      cache.put(cacheKey, content, DEFAULT_CACHE_DURATION);
    }
    return content || EMPTY;
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
 * @param {number} temperature [OPTIONAL] The randomness of the response. Lower values are more deterministic. Default is 0.0 but ChatGPT web default is 0.7.
 * @return {string|Array<Array<string>>} The response from ChatGPT.
 * @customfunction
 */
function CHATGPT(
  prompt: SpreadsheetInput<string>,
  model = "gpt-3.5-turbo",
  maxTokens = DEFAULT_MAX_TOKENS,
  temperature = DEFAULT_TEMPERATURE,
): SpreadsheetInput<string> {
  const apiKey = PROPERTY_API_KEY_GET();

  if (Array.isArray(prompt)) {
    return prompt.map((row) => {
      return row.map((cell) => {
        return REQUEST_COMPLETIONS(
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
  return REQUEST_COMPLETIONS(
    apiKey,
    SYSTEM_PROMPT,
    prompt,
    model,
    maxTokens,
    temperature,
  );
}

/**
 * Custom function to call ChatGPT-3 API. This is the default and most cost-effective model.
 * Example: =CHATGPT3("Summarize the plot of the movie: " & A1)
 *
 * @param {string|Array<Array<string>>} prompt The prompt to send to ChatGPT
 * @param {number} maxTokens [OPTIONAL] The maximum number of tokens to return. Default is 150, which is a short response. ChatGPT web default is 4096.
 * @param {number} temperature [OPTIONAL] The randomness of the response. Lower values are more deterministic. Default is 0.0 but ChatGPT web default is 0.7.
 * @return {string|Array<Array<string>>} The response from ChatGPT
 * @customfunction
 */
function CHATGPT3(
  prompt: SpreadsheetInput<string>,
  maxTokens = DEFAULT_MAX_TOKENS,
  temperature = DEFAULT_TEMPERATURE,
): SpreadsheetInput<string> {
  return CHATGPT(prompt, "gpt-3.5-turbo", maxTokens, temperature);
}

/**
 * Custom function to call ChatGPT-4 API. This is the latest and most powerful model.
 * Example: =CHATGPT4("Categorize the following text into 'positive' or 'negative': " & A1)
 *
 * @param {string|Array<Array<string>>} prompt The prompt to send to ChatGPT
 * @param {number} maxTokens [OPTIONAL] The maximum number of tokens to return. Default is 150, which is a short response. ChatGPT web default is 4096.
 * @param {number} temperature [OPTIONAL] The randomness of the response. Lower values are more deterministic. Default is 0.0 but ChatGPT web default is 0.7.
 * @return {string|Array<Array<string>>} The response from ChatGPT
 * @customfunction
 */
function CHATGPT4(
  prompt: SpreadsheetInput<string>,
  maxTokens = DEFAULT_MAX_TOKENS,
  temperature = DEFAULT_TEMPERATURE,
): SpreadsheetInput<string> {
  return CHATGPT(prompt, "gpt-4", maxTokens, temperature);
}

/**
 * Custom function to set the OpenAI API key. Get yours at: https://platform.openai.com/api-keys
 *
 * @param {string} apiKey The OpenAI API key to save
 * @return {string} Confirmation message
 * @customfunction
 */
function CHATGPTKEY(apiKey: string): string {
  PROPERTY_API_KEY_SET(apiKey);

  if (!apiKey) {
    return "🚮 API Key removed from user settings.";
  }

  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "get",
    contentType: MIME_JSON,
    headers: {
      Accept: MIME_JSON,
      Authorization: `Bearer ${apiKey}`,
    },
  };
  const response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/models",
    options,
  );
  const json = response.getContentText();
  const data = JSON.parse(json) as ModelsPage;
  if (!Array.isArray(data["data"]) || data["data"].length === 0) {
    return "❌ API Key is invalid or failed to connect.";
  }
  return "✅ API Key saved successfully.";
}

/**
 * Custom function to see available models from OpenAI.
 * Example: =CHATGPTMODELS()
 * @return {Array<Array<string>>} The list of available models
 * @customfunction
 */
function CHATGPTMODELS(): Array<Array<string>> {
  const apiKey = PROPERTY_API_KEY_GET();

  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "get",
    contentType: MIME_JSON,
    headers: {
      Accept: MIME_JSON,
      Authorization: `Bearer ${apiKey}`,
    },
  };
  const response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/models",
    options,
  );
  const json = response.getContentText();
  const data = JSON.parse(json) as ModelsPage;
  if (!Array.isArray(data["data"]) || data["data"].length === 0) {
    return [["🌵 No models available"]];
  }
  return data["data"].map((model) => {
    return [model["id"]];
  });
}

function PROPERTY_API_KEY_SET(apiKey: string | null | undefined) {
  const properties = PropertiesService.getUserProperties();
  if (apiKey === null || apiKey === undefined) {
    properties.deleteProperty(PROPERTY_KEY_OPENAPI);
  } else {
    properties.setProperty(PROPERTY_KEY_OPENAPI, apiKey);
  }
}

function PROPERTY_API_KEY_GET(): string {
  const properties = PropertiesService.getUserProperties();
  const apiKey = properties.getProperty(PROPERTY_KEY_OPENAPI) || OPENAI_API_KEY;
  if (!apiKey) {
    throw new Error(
      'Use =CHATGPTKEY("YOUR_API_KEY") first. Get it from https://platform.openai.com/api-keys',
    );
  }
  return apiKey;
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

function HASH(
  algorithm: typeof DigestAlgorithm,
  ...args: (string | number)[]
): string {
  const input = args.join("");
  const hash = Utilities.computeDigest(algorithm, input)
    .map((byte) => {
      const v = (byte & 0xff).toString(16);
      return v.length === 1 ? `0${v}` : v;
    })
    .join("");
  return hash;
}
function HASH_SHA1(...args: (string | number)[]): string {
  return HASH(Utilities.DigestAlgorithm.SHA_1, ...args);
}

function GET_CACHE() {
  // TODO: not sure which one is the best cache to use.
  return (
    CacheService.getDocumentCache() ||
    CacheService.getScriptCache() ||
    CacheService.getUserCache()
  );
}
