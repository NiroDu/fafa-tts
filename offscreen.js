let currentAudio = null;

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.type === "play-audio") {
    playAudio(message.dataUrl);
    sendResponse({ ok: true });
    return;
  }

  if (message?.type === "stop-audio") {
    stopAudio();
    sendResponse({ ok: true });
  }
});

function playAudio(dataUrl) {
  stopAudio();

  const audio = new Audio();
  audio.src = dataUrl;
  audio.onplay = () => {
    chrome.runtime.sendMessage({ type: "play-started" });
  };
  audio.onended = () => {
    if (currentAudio === audio) {
      currentAudio = null;
    }
    chrome.runtime.sendMessage({ type: "play-ended" });
  };
  audio.onerror = (event) => {
    console.error("fafa-tts audio error", event);
    if (currentAudio === audio) {
      currentAudio = null;
    }
    chrome.runtime.sendMessage({ type: "play-ended" });
  };

  currentAudio = audio;
  audio.play().catch((error) => {
    console.error("fafa-tts audio play failed", error);
    if (currentAudio === audio) {
      currentAudio = null;
    }
    chrome.runtime.sendMessage({ type: "play-ended" });
  });
}

function stopAudio() {
  if (!currentAudio) {
    return;
  }

  currentAudio.pause();
  currentAudio = null;
  chrome.runtime.sendMessage({ type: "play-ended" });
}
