function setStatus(message) {
  document.getElementById("status").textContent = message;
}

document.getElementById("fill").addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  if (!tab?.id) {
    setStatus("No active tab found.");
    return;
  }

  chrome.tabs.sendMessage(tab.id, { type: "FILL_BLOOMBERG_FORM" }, (response) => {
    if (chrome.runtime.lastError) {
      setStatus(chrome.runtime.lastError.message);
      return;
    }

    setStatus(response?.message || "Done.");
  });
});
