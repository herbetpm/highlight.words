
chrome.storage.local.get("highlightWords", ({ highlightWords }) => {
    highlightWords = highlightWords || [];
    const regex = new RegExp(`\\b(${highlightWords.join("|")})\\b`, "gi");

    const observer = new MutationObserver(() => {
        const captions = document.querySelectorAll("body *");
        captions.forEach((caption) => {
            caption.innerHTML = caption.innerHTML.replace(regex, (match) => {
                return `<mark>${match}</mark>`;
            });
        });
    });

    observer.observe(document.body, {
        childList: true,
        subtree: true
    });
});

document.addEventListener("click", (event) => {
    if (event.target.tagName === "SPAN" || event.target.tagName === "P" || event.target.tagName === "DIV") {
        const clickedWord = event.target.innerText.toLowerCase().trim();
        chrome.storage.local.get("highlightWords", ({ highlightWords }) => {
            highlightWords = highlightWords || [];
            if (!highlightWords.includes(clickedWord)) {
                highlightWords.push(clickedWord);
                highlightWords.sort();
                chrome.storage.local.set({ highlightWords }, () => {
                    chrome.runtime.sendMessage({ action: "saveWordsToOneDrive", words: highlightWords });
                });
            }
        });
    }
});
