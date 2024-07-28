
document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("manageWordsButton").addEventListener("click", () => {
        chrome.runtime.openOptionsPage();
    });

    document.getElementById("newWord").addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
            const newWord = event.target.value.trim().toLowerCase();
            if (newWord) {
                chrome.storage.local.get("highlightWords", ({ highlightWords }) => {
                    highlightWords = highlightWords || [];
                    if (!highlightWords.includes(newWord)) {
                        highlightWords.push(newWord);
                        highlightWords.sort();
                        chrome.storage.local.set({ highlightWords }, () => {
                            chrome.runtime.sendMessage({ action: "saveWordsToOneDrive", words: highlightWords });
                        });
                    }
                });
                event.target.value = '';
            }
        }
    });
});
