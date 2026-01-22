const DEFAULT_SETTINGS = {
  apiKey: "",
  region: "",
  rate: 1,
  defaultVoice: "zh-CN-XiaoxiaoNeural",
  voiceMap: {
    zh: "zh-CN-XiaoxiaoNeural",
    en: "en-US-JennyNeural",
  },
  outputFormat: "audio-24khz-48kbitrate-mono-mp3",
};

const LANGUAGE_ALIASES = {
  yue: "zh",
  lzh: "zh",
  wuu: "zh",
  nan: "zh",
  hak: "zh",
  cmn: "zh",
};

const LANGUAGE_LABELS = {
  zh: "中文",
  en: "英语",
  ja: "日语",
  ko: "韩语",
  fr: "法语",
  de: "德语",
  es: "西班牙语",
  it: "意大利语",
  ru: "俄语",
  pt: "葡萄牙语",
  th: "泰语",
  vi: "越南语",
  id: "印尼语",
  ms: "马来语",
  ar: "阿拉伯语",
  hi: "印地语",
  tr: "土耳其语",
  pl: "波兰语",
  nl: "荷兰语",
  sv: "瑞典语",
  no: "挪威语",
  da: "丹麦语",
  fi: "芬兰语",
  cs: "捷克语",
  hu: "匈牙利语",
  ro: "罗马尼亚语",
  el: "希腊语",
  uk: "乌克兰语",
  he: "希伯来语",
};

const elements = {
  apiKey: document.getElementById("apiKey"),
  region: document.getElementById("region"),
  rateRange: document.getElementById("rateRange"),
  rateValue: document.getElementById("rateValue"),
  refreshVoices: document.getElementById("refreshVoices"),
  voicesStatus: document.getElementById("voicesStatus"),
  defaultVoice: document.getElementById("defaultVoice"),
  voiceSplit: document.getElementById("voiceSplit"),
  voiceList: document.getElementById("voiceList"),
  voiceDetail: document.getElementById("voiceDetail"),
  testConnection: document.getElementById("testConnection"),
  testStatus: document.getElementById("testStatus"),
  cacheCount: document.getElementById("cacheCount"),
  toggleCache: document.getElementById("toggleCache"),
  cachePanel: document.getElementById("cachePanel"),
  cacheList: document.getElementById("cacheList"),
  clearCache: document.getElementById("clearCache"),
  save: document.getElementById("save"),
  saveStatus: document.getElementById("saveStatus"),
};

let voiceGroups = [];
let draftVoiceMap = {};
let activeLanguage = null;
let previewAudio = null;
let lastSavedSettings = null;

init();

async function init() {
  const stored = await chrome.storage.local.get([
    "settings",
    "audioCache",
    "audioCacheOrder",
    "voicesCache",
  ]);
  const settings = {
    ...DEFAULT_SETTINGS,
    ...(stored.settings || {}),
    voiceMap: {
      ...DEFAULT_SETTINGS.voiceMap,
      ...((stored.settings || {}).voiceMap || {}),
    },
  };

  lastSavedSettings = settings;
  elements.apiKey.value = settings.apiKey || "";
  elements.region.value = settings.region || "";
  elements.rateRange.value = String(clampRate(settings.rate));
  updateRateDisplay(Number(elements.rateRange.value));

  updateCacheCount(stored.audioCache || {});

  elements.rateRange.addEventListener("input", () => {
    updateRateDisplay(Number(elements.rateRange.value));
  });

  elements.refreshVoices.addEventListener("click", () => {
    loadVoiceCatalog({ force: true });
  });

  elements.defaultVoice.addEventListener("change", () => {
    renderVoiceList();
    renderVoiceDetail();
  });

  elements.save.addEventListener("click", saveSettings);
  elements.testConnection.addEventListener("click", testConnection);
  elements.clearCache.addEventListener("click", clearCache);
  elements.toggleCache.addEventListener("click", toggleCachePanel);

  if (settings.apiKey && settings.region) {
    await loadVoiceCatalog({ cached: stored.voicesCache, settings });
  } else {
    renderVoiceSelectors([], settings);
  }

  if (!elements.cachePanel.classList.contains("hidden")) {
    renderCacheList(stored.audioCache || {}, stored.audioCacheOrder || []);
  }
}

function updateRateDisplay(value) {
  const normalized = clampRate(value);
  elements.rateValue.textContent = `${normalized.toFixed(1)}x`;
}

function clampRate(rate) {
  const parsed = Number(rate);
  if (!Number.isFinite(parsed)) {
    return 1;
  }
  return Math.min(2, Math.max(0.5, parsed));
}

async function loadVoiceCatalog({
  force = false,
  cached = null,
  settings = null,
} = {}) {
  const config = getConfigFromForm();
  if (!config.apiKey || !config.region) {
    setStatus(elements.voicesStatus, "请先填写 API Key 和区域", false);
    renderVoiceSelectors([], settings || getSettingsFromForm());
    return;
  }

  if (
    !force &&
    cached &&
    cached.region === config.region &&
    Array.isArray(cached.voices)
  ) {
    renderVoiceSelectors(cached.voices, settings || getSettingsFromForm());
    setStatus(elements.voicesStatus, "已加载缓存声音列表", true);
    return;
  }

  setStatus(elements.voicesStatus, "加载声音列表...", null);

  try {
    const voices = await fetchVoiceList(config.apiKey, config.region);
    await chrome.storage.local.set({
      voicesCache: {
        region: config.region,
        voices,
        fetchedAt: Date.now(),
      },
    });
    renderVoiceSelectors(voices, settings || getSettingsFromForm());
    setStatus(elements.voicesStatus, `已加载 ${voices.length} 个声音`, true);
  } catch (error) {
    console.error(error);
    setStatus(elements.voicesStatus, "声音列表加载失败", false);
    renderVoiceSelectors([], settings || getSettingsFromForm());
  }
}

async function fetchVoiceList(apiKey, region) {
  const endpoint = `https://${region}.tts.speech.microsoft.com/cognitiveservices/voices/list`;
  const response = await fetch(endpoint, {
    headers: {
      "Ocp-Apim-Subscription-Key": apiKey,
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Azure voices list error ${response.status}: ${errorText}`);
  }

  const voices = await response.json();
  return Array.isArray(voices) ? voices : [];
}

function renderVoiceSelectors(voices, settings) {
  voiceGroups = buildLanguageGroups(voices);
  draftVoiceMap = normalizeVoiceMap(settings.voiceMap || {}, voiceGroups);
  renderDefaultVoice(voices, settings.defaultVoice);

  if (
    !activeLanguage ||
    !voiceGroups.some((group) => group.key === activeLanguage)
  ) {
    activeLanguage = voiceGroups[0]?.key || null;
  }

  if (voiceGroups.length === 0) {
    elements.voiceSplit.classList.remove("expanded");
  }

  renderVoiceList();
  renderVoiceDetail();
}

function renderVoiceList() {
  elements.voiceList.innerHTML = "";

  if (voiceGroups.length === 0) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "暂无声音列表，刷新后可选择。";
    elements.voiceList.appendChild(empty);
    elements.voiceDetail.innerHTML = "";
    return;
  }

  voiceGroups.forEach((group) => {
    const item = document.createElement("button");
    item.type = "button";
    item.className = "list-item";
    item.dataset.languageKey = group.key;
    if (group.key === activeLanguage) {
      item.classList.add("active");
    }

    const title = document.createElement("div");
    title.className = "list-title";
    title.textContent = group.label;

    const subtitle = document.createElement("div");
    subtitle.className = "list-subtitle";
    subtitle.textContent = `${getVoiceLabelForGroup(group)} · ${group.voices.length} 个声音`;

    item.append(title, subtitle);
    item.addEventListener("click", () => {
      activeLanguage = group.key;
      elements.voiceSplit.classList.add("expanded");
      renderVoiceList();
      renderVoiceDetail();
    });

    elements.voiceList.appendChild(item);
  });
}

function renderVoiceDetail() {
  elements.voiceDetail.innerHTML = "";

  if (!activeLanguage) {
    return;
  }

  const group = voiceGroups.find((item) => item.key === activeLanguage);
  if (!group) {
    return;
  }

  const detail = document.createElement("div");
  detail.className = "detail-card";

  const title = document.createElement("div");
  title.className = "list-title";
  title.textContent = group.label;

  const select = document.createElement("select");
  group.voices.forEach((voice) => {
    const option = document.createElement("option");
    option.value = voice.ShortName;
    option.textContent = formatVoiceLabel(voice, true);
    select.appendChild(option);
  });

  const preferredVoice = getPreferredVoiceForGroup(group);
  if (preferredVoice) {
    select.value = preferredVoice;
  }

  select.addEventListener("change", () => {
    draftVoiceMap[group.key] = select.value;
    renderVoiceList();
  });

  const actions = document.createElement("div");
  actions.className = "detail-actions";

  const previewButton = document.createElement("button");
  previewButton.type = "button";
  previewButton.className = "secondary";
  previewButton.textContent = "试听";
  previewButton.addEventListener("click", () => {
    previewVoice(select.value, group.sampleLocale);
  });

  actions.append(previewButton);
  detail.append(title, select, actions);
  elements.voiceDetail.appendChild(detail);
}

function buildLanguageGroups(voices) {
  const groups = new Map();
  voices.forEach((voice) => {
    if (!voice.Locale || !voice.ShortName) {
      return;
    }

    const base = voice.Locale.split("-")[0].toLowerCase();
    const key = normalizeLanguageKey(base);
    if (!groups.has(key)) {
      groups.set(key, {
        key,
        label: getLanguageLabel(key),
        voices: [],
        sampleLocale: voice.Locale,
      });
    }

    groups.get(key).voices.push(voice);
  });

  return Array.from(groups.values())
    .map((group) => ({
      ...group,
      voices: group.voices.sort((a, b) =>
        a.ShortName.localeCompare(b.ShortName),
      ),
    }))
    .sort((a, b) => compareLanguagePriority(a.key, b.key, a.label, b.label));
}

function compareLanguagePriority(aKey, bKey, aLabel, bLabel) {
  const priority = ["en", "zh", "ja"];
  const aIndex = priority.indexOf(aKey);
  const bIndex = priority.indexOf(bKey);
  if (aIndex !== -1 || bIndex !== -1) {
    if (aIndex === -1) {
      return 1;
    }
    if (bIndex === -1) {
      return -1;
    }
    return aIndex - bIndex;
  }
  return aLabel.localeCompare(bLabel);
}

function normalizeLanguageKey(base) {
  return LANGUAGE_ALIASES[base] || base;
}

function getLanguageLabel(key) {
  return LANGUAGE_LABELS[key] || key.toUpperCase();
}

function normalizeVoiceMap(voiceMap, groups) {
  const normalized = {};
  Object.entries(voiceMap || {}).forEach(([key, value]) => {
    if (!value) {
      return;
    }
    const base = normalizeLanguageKey(String(key).split("-")[0].toLowerCase());
    if (!normalized[base]) {
      normalized[base] = value;
    }
  });

  groups.forEach((group) => {
    const selected = normalized[group.key];
    if (
      selected &&
      !group.voices.some((voice) => voice.ShortName === selected)
    ) {
      delete normalized[group.key];
    }
  });

  return normalized;
}

function getPreferredVoiceForGroup(group) {
  const storedVoice = draftVoiceMap[group.key];
  if (
    storedVoice &&
    group.voices.some((voice) => voice.ShortName === storedVoice)
  ) {
    return storedVoice;
  }

  const defaultVoice = getDefaultVoiceValue();
  if (
    defaultVoice &&
    group.voices.some((voice) => voice.ShortName === defaultVoice)
  ) {
    return defaultVoice;
  }

  return group.voices[0]?.ShortName || "";
}

function getVoiceLabelForGroup(group) {
  const voiceName = getPreferredVoiceForGroup(group);
  if (!voiceName) {
    return "未选择";
  }

  const voice = group.voices.find((item) => item.ShortName === voiceName);
  return voice ? formatVoiceLabel(voice) : voiceName;
}

function renderDefaultVoice(voices, selected) {
  elements.defaultVoice.innerHTML = "";

  if (!voices || voices.length === 0) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "请先刷新声音列表";
    elements.defaultVoice.appendChild(option);
    return;
  }

  voices.forEach((voice) => {
    if (!voice.ShortName) {
      return;
    }
    const option = document.createElement("option");
    option.value = voice.ShortName;
    option.textContent = formatVoiceLabel(voice, true);
    elements.defaultVoice.appendChild(option);
  });

  if (selected) {
    elements.defaultVoice.value = selected;
  }
}

function formatVoiceLabel(voice, includeLocale = false) {
  const gender = voice.Gender ? ` ${voice.Gender}` : "";
  if (includeLocale) {
    const localeName = voice.LocaleName || voice.Locale || "";
    const localeLabel = localeName ? ` · ${localeName}` : "";
    return `${voice.ShortName}${gender}${localeLabel}`.trim();
  }
  return `${voice.ShortName}${gender}`.trim();
}

async function previewVoice(voiceName, locale) {
  const config = getConfigFromForm();
  if (!config.apiKey || !config.region) {
    setStatus(elements.voicesStatus, "请先填写 API Key 和区域", false);
    return;
  }

  if (!voiceName) {
    setStatus(elements.voicesStatus, "未选择声音", false);
    return;
  }

  setStatus(elements.voicesStatus, "试听中...", null);

  try {
    const rateMultiplier = clampRate(elements.rateRange.value);
    const dataUrl = await fetchTts({
      text: sampleTextForLocale(locale),
      voice: voiceName,
      rate: rateToProsody(rateMultiplier),
      region: config.region,
      apiKey: config.apiKey,
      outputFormat: DEFAULT_SETTINGS.outputFormat,
    });
    playPreviewAudio(dataUrl);
    setStatus(elements.voicesStatus, "试听完成", true);
  } catch (error) {
    console.error(error);
    setStatus(elements.voicesStatus, "试听失败", false);
  }
}

async function saveSettings() {
  const settings = getSettingsFromForm();
  await chrome.storage.local.set({ settings });
  lastSavedSettings = settings;
  setStatus(elements.saveStatus, "已保存", true);
}

function getConfigFromForm() {
  return {
    apiKey: elements.apiKey.value.trim(),
    region: elements.region.value.trim(),
  };
}

function getDefaultVoiceValue() {
  return (
    elements.defaultVoice.value ||
    lastSavedSettings?.defaultVoice ||
    DEFAULT_SETTINGS.defaultVoice
  );
}

function getSettingsFromForm() {
  const voiceMap =
    Object.keys(draftVoiceMap).length > 0
      ? draftVoiceMap
      : lastSavedSettings?.voiceMap || DEFAULT_SETTINGS.voiceMap;

  return {
    ...DEFAULT_SETTINGS,
    apiKey: elements.apiKey.value.trim(),
    region: elements.region.value.trim(),
    rate: clampRate(elements.rateRange.value),
    defaultVoice: getDefaultVoiceValue(),
    voiceMap,
  };
}

async function testConnection() {
  const settings = getSettingsFromForm();
  if (!settings.apiKey || !settings.region) {
    setStatus(elements.testStatus, "请先填写 API Key 和区域", false);
    return;
  }

  if (!settings.defaultVoice) {
    setStatus(elements.testStatus, "请先刷新并选择默认声音", false);
    return;
  }

  setStatus(elements.testStatus, "请求中...", null);

  try {
    await fetchTts({
      text: sampleTextForLocale(localeFromVoice(settings.defaultVoice)),
      voice: settings.defaultVoice,
      rate: rateToProsody(clampRate(settings.rate)),
      region: settings.region,
      apiKey: settings.apiKey,
      outputFormat: settings.outputFormat,
    });
    setStatus(elements.testStatus, "连接成功", true);
  } catch (error) {
    console.error(error);
    setStatus(elements.testStatus, "连接失败，请检查 Key/区域/声音", false);
  }
}

async function clearCache() {
  await chrome.storage.local.remove(["audioCache", "audioCacheOrder"]);
  updateCacheCount({});
  elements.cacheList.innerHTML = "";
  setStatus(elements.saveStatus, "已清空缓存", true);
}

function toggleCachePanel() {
  elements.cachePanel.classList.toggle("hidden");
  const isHidden = elements.cachePanel.classList.contains("hidden");
  elements.toggleCache.textContent = isHidden ? "查看缓存" : "收起缓存";

  if (!isHidden) {
    refreshCacheList();
  }
}

async function refreshCacheList() {
  const stored = await chrome.storage.local.get([
    "audioCache",
    "audioCacheOrder",
  ]);
  updateCacheCount(stored.audioCache || {});
  renderCacheList(stored.audioCache || {}, stored.audioCacheOrder || []);
}

function renderCacheList(audioCache, cacheOrder) {
  elements.cacheList.innerHTML = "";

  const keys =
    cacheOrder && cacheOrder.length > 0
      ? [...cacheOrder].reverse()
      : Object.keys(audioCache || {});

  if (keys.length === 0) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "暂无缓存内容";
    elements.cacheList.appendChild(empty);
    return;
  }

  keys.forEach((key) => {
    const entry = audioCache[key];
    const dataUrl = typeof entry === "string" ? entry : entry?.dataUrl;
    const text = typeof entry === "string" ? "" : entry?.text;
    const voice = typeof entry === "string" ? "" : entry?.voice;
    const rate = typeof entry === "string" ? null : entry?.rate;
    const createdAt = typeof entry === "string" ? null : entry?.createdAt;

    const card = document.createElement("div");
    card.className = "detail-card";

    const body = document.createElement("div");
    body.className = "detail-text";
    body.textContent = text || "旧缓存条目（无文本内容）";

    const metaRow = document.createElement("div");
    metaRow.className = "detail-row";

    const meta = document.createElement("div");
    meta.className = "detail-meta";
    meta.textContent = buildCacheMeta({ voice, rate, createdAt }) || "";

    const actions = document.createElement("div");
    actions.className = "detail-actions";

    const playButton = document.createElement("button");
    playButton.type = "button";
    playButton.className = "secondary";
    playButton.textContent = "播放";
    playButton.disabled = !dataUrl;
    playButton.addEventListener("click", () => {
      if (dataUrl) {
        playPreviewAudio(dataUrl);
      }
    });

    const downloadButton = document.createElement("button");
    downloadButton.type = "button";
    downloadButton.className = "secondary";
    downloadButton.textContent = "下载";
    downloadButton.disabled = !dataUrl;
    downloadButton.addEventListener("click", () => {
      if (dataUrl) {
        downloadDataUrl(dataUrl, buildDownloadName(entry));
      }
    });

    actions.append(playButton, downloadButton);
    metaRow.append(actions, meta);
    card.append(body, metaRow);
    elements.cacheList.appendChild(card);
  });
}

function buildDownloadName(entry) {
  const createdAt = typeof entry === "string" ? null : entry?.createdAt;
  const timestamp = createdAt ? new Date(createdAt) : new Date();
  const iso = timestamp.toISOString().replace(/[:.]/g, "-");
  return `fafa-tts-${iso}.mp3`;
}

function downloadDataUrl(dataUrl, filename) {
  const link = document.createElement("a");
  link.href = dataUrl;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
}

function buildCacheMeta({ voice, rate, createdAt }) {
  const parts = [];
  if (voice) {
    parts.push(voice);
  }
  if (rate && Number.isFinite(rate)) {
    parts.push(`语速 ${rate.toFixed(1)}x`);
  }
  if (createdAt) {
    parts.push(new Date(createdAt).toLocaleString());
  }
  return parts.join(" · ") || "";
}

function updateCacheCount(audioCache) {
  const count = Object.keys(audioCache || {}).length;
  elements.cacheCount.textContent = String(count);
}

function setStatus(element, message, ok) {
  element.textContent = message;
  if (ok === true) {
    element.style.color = "#107c10";
  } else if (ok === false) {
    element.style.color = "#b91c1c";
  } else {
    element.style.color = "#5f5c57";
  }
}

function rateToProsody(rateMultiplier) {
  const percent = Math.round((rateMultiplier - 1) * 100);
  return `${percent}%`;
}

function localeFromVoice(voice) {
  const match = String(voice || "").match(/^[a-z]{2}-[A-Z]{2}/);
  return match ? match[0] : "en-US";
}

function sampleTextForLocale(locale) {
  if (!locale) {
    return "Hello, this is a voice preview.";
  }

  const prefix = locale.split("-")[0];
  switch (prefix) {
    case "zh":
    case "yue":
    case "lzh":
    case "wuu":
    case "nan":
    case "hak":
    case "cmn":
      return "你好，这是一段试听音频。";
    case "ja":
      return "こんにちは、これは音声プレビューです。";
    case "ko":
      return "안녕하세요, 음성 미리듣기입니다.";
    case "fr":
      return "Bonjour, ceci est un aperçu vocal.";
    case "de":
      return "Hallo, dies ist eine Sprachvorschau.";
    case "es":
      return "Hola, esta es una vista previa de voz.";
    case "it":
      return "Ciao, questa è un'anteprima vocale.";
    case "ru":
      return "Здравствуйте, это предварительный просмотр голоса.";
    default:
      return "Hello, this is a voice preview.";
  }
}

async function fetchTts({ text, voice, rate, region, apiKey, outputFormat }) {
  const locale = localeFromVoice(voice);
  const ssml =
    `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n` +
    `<speak version=\"1.0\" xml:lang=\"${locale}\">` +
    `<voice name=\"${voice}\">` +
    `<prosody rate=\"${rate}\">${escapeForSsml(text)}</prosody>` +
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

function escapeForSsml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function playPreviewAudio(dataUrl) {
  if (previewAudio) {
    previewAudio.pause();
    previewAudio = null;
  }

  const audio = new Audio();
  audio.src = dataUrl;
  audio.onended = () => {
    if (previewAudio === audio) {
      previewAudio = null;
    }
  };
  audio.onerror = (event) => {
    console.error("fafa-tts preview error", event);
    if (previewAudio === audio) {
      previewAudio = null;
    }
  };

  previewAudio = audio;
  audio.play().catch((error) => {
    console.error("fafa-tts preview play failed", error);
    if (previewAudio === audio) {
      previewAudio = null;
    }
  });
}
