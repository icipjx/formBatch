/**
 * FormBatcher MVP background service worker
 * - Guard sidePanel API for older Chrome / unsupported environments.
 */
function tryEnableOpenOnClick() {
  try {
    // Some Chrome versions / Chromium builds don't expose chrome.sidePanel
    if (chrome.sidePanel && chrome.sidePanel.setPanelBehavior) {
      chrome.sidePanel
        .setPanelBehavior({ openPanelOnActionClick: true })
        .catch(() => {});
    }
  } catch (e) {
    // ignore
  }
}

chrome.runtime.onInstalled.addListener(() => {
  tryEnableOpenOnClick();
});

chrome.runtime.onStartup.addListener(() => {
  tryEnableOpenOnClick();
});

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg?._fbForwarded) return;
  if (msg?.type === "FB_PICK_RESULT" || msg?.type === "FB_PICK_CANCELLED" || msg?.type === "FB_RUN_EVENT") {
    chrome.runtime.sendMessage({ ...msg, _fbForwarded: true });
  }
});

chrome.action.onClicked.addListener((tab) => {
  try {
    if (chrome.sidePanel && chrome.sidePanel.open && tab && tab.id !== undefined) {
      chrome.sidePanel.open({ tabId: tab.id }).catch(() => {});
      return;
    }
    tryEnableOpenOnClick();
  } catch (e) {
    // ignore
  }
});
