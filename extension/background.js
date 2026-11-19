// Loads only the active scene's CSS into the Word frame that asked for it.
chrome.runtime.onMessage.addListener((m, s) => {
  if (m.css && /^[a-z0-9]+$/.test(m.css)) chrome.scripting.insertCSS({ target: { tabId: s.tab.id, frameIds: [s.frameId] }, files: [`scenes/${m.css}.css`] });
});
chrome.runtime.onInstalled.addListener(({ reason }) => {
  if (reason == "install") chrome.tabs.create({ url: chrome.runtime.getURL("options.html#file") });
});
