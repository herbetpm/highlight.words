
const msalConfig = {
    auth: {
        clientId: "e43923d1-c400-4588-843b-3cb5eb99107a",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "https://herbetpm.github.io/highlight.words/options.html"
    },
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

document.getElementById("loginButton").addEventListener("click", () => {
    signIn().then(() => {
        console.log("Logged in successfully");
        syncWords();
    }).catch(error => {
        console.error("Login failed", error);
    });
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
                        updateWordList(highlightWords);
                        saveWordsToOneDrive(highlightWords);
                    });
                }
            });
            event.target.value = '';
        }
    }
});

document.addEventListener("DOMContentLoaded", () => {
    chrome.storage.local.get("highlightWords", ({ highlightWords }) => {
        if (chrome.runtime.lastError) {
            console.error(chrome.runtime.lastError);
            return;
        }
        highlightWords = highlightWords || [];
        updateWordList(highlightWords);
    });
});

async function signIn() {
    const loginRequest = {
        scopes: ["User.Read", "Files.ReadWrite", "offline_access"]
    };

    try {
        const loginResponse = await msalInstance.loginPopup(loginRequest);
        console.log("id_token acquired at: " + new Date().toString());
        console.log(loginResponse);

        const tokenRequest = {
            scopes: ["User.Read", "Files.ReadWrite"],
            account: loginResponse.account
        };

        const tokenResponse = await msalInstance.acquireTokenSilent(tokenRequest);
        console.log("access_token acquired at: " + new Date().toString());
        console.log(tokenResponse);
        return tokenResponse.accessToken;
    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            return msalInstance.acquireTokenPopup(loginRequest).then(tokenResponse => {
                return tokenResponse.accessToken;
            }).catch(innerError => {
                console.error(innerError);
            });
        } else {
            console.error(error);
        }
    }
}

async function saveWordsToOneDrive(words) {
    const accessToken = await signIn();
    const content = JSON.stringify(words);

    const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/highlightWords.json:/content`;
    try {
        await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
                Authorization: `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: content
        });
        console.log("Words saved to OneDrive.");
    } catch (error) {
        console.error("Error saving words to OneDrive:", error);
    }
}

async function loadWordsFromOneDrive() {
    const accessToken = await signIn();
    const downloadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/highlightWords.json:/content`;

    try {
        const response = await fetch(downloadUrl, {
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        });

        if (response.ok) {
            const data = await response.json();
            console.log("Words loaded from OneDrive:", data);
            return data;
        } else if (response.status === 404) {
            console.log("File not found, returning empty array");
            return [];
        } else {
            console.error("Error loading words from OneDrive:", response.status, response.statusText);
            return [];
        }
    } catch (error) {
        console.error("Error loading words from OneDrive:", error);
        return [];
    }
}

async function syncWords() {
    const words = await loadWordsFromOneDrive();
    if (!Array.isArray(words)) {
        console.log("No words found in OneDrive, initializing empty array");
        words = [];
    }
    chrome.storage.local.set({ highlightWords: words }, () => {
        console.log("Words synced to local storage.");
        updateWordList(words);
    });
}

function updateWordList(words) {
    const wordList = document.getElementById('wordList');
    wordList.innerHTML = '';
    words.forEach(word => {
        const li = document.createElement('li');
        li.textContent = word;
        const deleteButton = document.createElement('button');
        deleteButton.textContent = 'Delete';
        deleteButton.onclick = () => {
            removeWord(word);
        };
        li.appendChild(deleteButton);
        wordList.appendChild(li);
    });
}

function removeWord(word) {
    chrome.storage.local.get('highlightWords', ({ highlightWords }) => {
        highlightWords = highlightWords.filter(w => w !== word);
        chrome.storage.local.set({ highlightWords }, () => {
            updateWordList(highlightWords);
            saveWordsToOneDrive(highlightWords);
        });
    });
}
