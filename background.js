const DEFAULT_SETTINGS = {
  apiKey: "",
  region: "",
  rate: 1,
  defaultVoice: "zh-CN-XiaoxiaoNeural",
  voiceMap: {
    zh: "zh-CN-XiaoxiaoNeural",
    en: "en-GB-OliverNeural",
  },
  outputFormat: "audio-24khz-48kbitrate-mono-mp3",
  cacheMaxEntries: 50,
};

const CACHE_KEYS = ["audioCache", "audioCacheOrder"];
const LANGUAGE_ALIASES = {
  yue: "zh",
  lzh: "zh",
  wuu: "zh",
  nan: "zh",
  hak: "zh",
  cmn: "zh",
};

let isPlaying = false;

chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.create({
    id: "fafa-tts-read",
    title: "朗读",
    contexts: ["selection"],
  });
  chrome.contextMenus.create({
    id: "fafa-tts-settings",
    title: "打开设置",
    contexts: ["action"],
  });
});

chrome.contextMenus.onClicked.addListener(async (info) => {
  if (info.menuItemId !== "fafa-tts-read") {
    if (info.menuItemId === "fafa-tts-settings") {
      chrome.runtime.openOptionsPage();
    }
    return;
  }

  const text = (info.selectionText || "").trim();
  if (!text) {
    return;
  }

  try {
    await speakText(text);
  } catch (error) {
    console.error("fafa-tts failed:", error);
  }
});

chrome.action.onClicked.addListener(async (tab) => {
  if (isPlaying) {
    await stopPlayback();
    return;
  }
  if (!tab?.id) {
    return;
  }

  const text = await getSelectionFromTab(tab.id);
  if (!text) {
    return;
  }

  try {
    await speakText(text);
  } catch (error) {
    console.error("fafa-tts failed:", error);
  }
});

chrome.runtime.onMessage.addListener((message) => {
  if (message?.type === "play-started") {
    setPlayingState(true);
  }
  if (message?.type === "play-ended") {
    setPlayingState(false);
  }
});

async function getSettings() {
  const stored = await chrome.storage.local.get(["settings"]);
  const settings = stored.settings || {};
  return {
    ...DEFAULT_SETTINGS,
    ...settings,
    voiceMap: {
      ...DEFAULT_SETTINGS.voiceMap,
      ...(settings.voiceMap || {}),
    },
  };
}

async function speakText(text) {
  const settings = await getSettings();
  if (!settings.apiKey || !settings.region) {
    console.warn("fafa-tts: missing apiKey/region");
    return;
  }

  const detected = await detectLanguage(text);
  const voice = pickVoice(settings, detected);
  const rateMultiplier = normalizeRateMultiplier(settings.rate);
  const prosodyRate = rateToProsody(rateMultiplier);
  const cacheKey = await buildCacheKey(
    text,
    voice,
    rateMultiplier,
    settings.region,
    settings.outputFormat,
  );

  const cacheStore = await chrome.storage.local.get(CACHE_KEYS);
  const audioCache = cacheStore.audioCache || {};
  if (audioCache[cacheKey]) {
    const cachedEntry = audioCache[cacheKey];
    const dataUrl =
      typeof cachedEntry === "string" ? cachedEntry : cachedEntry?.dataUrl;
    if (dataUrl) {
      await playAudio(dataUrl);
      return;
    }
  }

  const dataUrl = await fetchTts({
    text,
    voice,
    rate: prosodyRate,
    region: settings.region,
    apiKey: settings.apiKey,
    outputFormat: settings.outputFormat,
    lang: detected?.language,
  });

  await saveToCache(
    cacheKey,
    {
      dataUrl,
      text,
      voice,
      rate: rateMultiplier,
      createdAt: Date.now(),
    },
    settings.cacheMaxEntries,
  );
  await playAudio(dataUrl);
}

async function detectLanguage(text) {
  try {
    const result = await chrome.i18n.detectLanguage(text);
    if (!result || !result.languages || result.languages.length === 0) {
      return null;
    }

    const [best] = result.languages.sort((a, b) => b.percentage - a.percentage);
    return best || null;
  } catch (error) {
    console.warn("fafa-tts: detectLanguage failed", error);
    return null;
  }
}

async function getSelectionFromTab(tabId) {
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId },
      func: () => (window.getSelection ? window.getSelection().toString() : ""),
    });
    const selection = results && results[0] && results[0].result;
    return String(selection || "").trim();
  } catch (error) {
    console.warn("fafa-tts: failed to read selection", error);
    return "";
  }
}

function pickVoice(settings, detected) {
  const voiceMap = settings.voiceMap || {};
  if (detected?.language) {
    const language = detected.language;
    if (voiceMap[language]) {
      return voiceMap[language];
    }
    const [base] = language.split("-");
    if (base && voiceMap[base]) {
      return voiceMap[base];
    }
    const normalizedBase = base ? normalizeLanguageKey(base) : "";
    if (normalizedBase && voiceMap[normalizedBase]) {
      return voiceMap[normalizedBase];
    }
    if (normalizedBase) {
      const defaultLocale = localeFromVoice(settings.defaultVoice || "");
      if (defaultLocale.startsWith(`${normalizedBase}-`)) {
        return settings.defaultVoice;
      }
      const keyMatch = Object.keys(voiceMap).find((key) =>
        key.startsWith(`${normalizedBase}-`),
      );
      if (keyMatch) {
        return voiceMap[keyMatch];
      }
    }
  }

  return (
    settings.defaultVoice ||
    Object.values(voiceMap)[0] ||
    DEFAULT_SETTINGS.defaultVoice
  );
}

function normalizeRateMultiplier(rate) {
  if (typeof rate === "number" && Number.isFinite(rate)) {
    return clampRate(rate);
  }

  const trimmed = String(rate || "").trim();
  if (!trimmed) {
    return 1;
  }

  if (trimmed.endsWith("%")) {
    const percent = Number(trimmed.slice(0, -1));
    if (Number.isFinite(percent)) {
      return clampRate(1 + percent / 100);
    }
  }

  const parsed = Number(trimmed);
  if (Number.isFinite(parsed)) {
    return clampRate(parsed);
  }

  return 1;
}

function clampRate(rate) {
  return Math.min(2, Math.max(0.5, rate));
}

function rateToProsody(rateMultiplier) {
  const percent = Math.round((rateMultiplier - 1) * 100);
  return `${percent}%`;
}

function localeFromVoice(voice) {
  const match = String(voice || "").match(/^[a-z]{2}-[A-Z]{2}/);
  return match ? match[0] : "en-US";
}

function normalizeLanguageKey(base) {
  return LANGUAGE_ALIASES[base] || base;
}

function escapeForSsml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

async function fetchTts({
  text,
  voice,
  rate,
  region,
  apiKey,
  outputFormat,
  lang,
}) {
  const locale = localeFromVoice(voice) || (lang ? lang : "en-US");
  const ssml =
    `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<speak version="1.0" xml:lang="${locale}">` +
    `<voice name="${voice}">` +
    `<prosody rate="${rate}">${escapeForSsml(text)}</prosody>` +
    `</voice>` +
    `</speak>`;

  const endpoint = `https://${region}.tts.speech.microsoft.com/cognitiveservices/v1`;
  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      "Ocp-Apim-Subscription-Key": apiKey,
      "Content-Type": "application/ssml+xml",
      "X-Microsoft-OutputFormat": outputFormat,
      "User-Agent": "fafa-tts",
    },
    body: ssml,
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Azure TTS error ${response.status}: ${errorText}`);
  }

  const arrayBuffer = await response.arrayBuffer();
  const base64 = arrayBufferToBase64(arrayBuffer);
  return `data:audio/mpeg;base64,${base64}`;
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (let i = 0; i < bytes.byteLength; i += 1) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

async function saveToCache(cacheKey, entry, maxEntries) {
  const store = await chrome.storage.local.get(CACHE_KEYS);
  const audioCache = store.audioCache || {};
  const audioCacheOrder = store.audioCacheOrder || [];

  if (!audioCache[cacheKey]) {
    audioCacheOrder.push(cacheKey);
  }

  audioCache[cacheKey] = entry;

  while (audioCacheOrder.length > maxEntries) {
    const oldest = audioCacheOrder.shift();
    if (oldest) {
      delete audioCache[oldest];
    }
  }

  await chrome.storage.local.set({ audioCache, audioCacheOrder });
}

async function buildCacheKey(text, voice, rateMultiplier, region, format) {
  const rateKey = Number.isFinite(rateMultiplier)
    ? rateMultiplier.toFixed(2)
    : "1.00";
  const data = `${region}|${voice}|${rateKey}|${format}|${text}`;
  const digest = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(data),
  );
  const hashArray = Array.from(new Uint8Array(digest));
  return hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function ensureOffscreenDocument() {
  const hasDoc = await chrome.offscreen.hasDocument();
  if (hasDoc) {
    return;
  }

  await chrome.offscreen.createDocument({
    url: "offscreen.html",
    reasons: ["AUDIO_PLAYBACK"],
    justification: "Play synthesized speech",
  });
}

async function playAudio(dataUrl) {
  await ensureOffscreenDocument();
  await chrome.runtime.sendMessage({
    type: "play-audio",
    dataUrl,
  });
  setPlayingState(true);
}

async function stopPlayback() {
  await ensureOffscreenDocument();
  await chrome.runtime.sendMessage({
    type: "stop-audio",
  });
  setPlayingState(false);
}

function setPlayingState(nextState) {
  if (isPlaying === nextState) {
    return;
  }
  isPlaying = nextState;
  chrome.action.setBadgeText({
    text: isPlaying ? "●" : "",
  });
  if (isPlaying) {
    chrome.action.setBadgeBackgroundColor({ color: "#d97706" });
  }
}
