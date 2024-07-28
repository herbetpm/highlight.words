
const msalConfig = {
    auth: {
        clientId: "e43923d1-c400-4588-843b-3cb5eb99107a",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "https://herbetpm.github.io/highlight.words/options.html"
    }
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

document.addEventListener("DOMContentLoaded", async () => {
    try {
        await handleRedirect();
        chrome.storage.local.get("highlightWords", ({ highlightWords }) => {
            if (chrome.runtime.lastError) {
                console.error(chrome.runtime.lastError);
                return;
            }
            highlightWords = highlightWords || [];
            updateWordList(highlightWords);
        });
    } catch (error) {
        console.error("Error during DOMContentLoaded:", error);
    }
});

async function handleRedirect() {
    try {
        const msalRedirectResponse = await msalInstance.handleRedirectPromise();
        
        if (msalRedirectResponse !== null) {
            console.log("Redirected back to application.");
            console.log("ID token acquired at: " + new Date().toString());
            console.log(msalRedirectResponse);
        }
    } catch (error) {
        console.error("Error handling redirect:", error);
    }
}

async function signIn() {
    const loginRequest = {
        scopes: ["User.Read", "Files.ReadWrite", "offline_access", "openid", "profile"]
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
                throw innerError;
            });
        } else {
            console.error(error);
            throw error;
        }
    }
}

async function saveWordsToOneDrive(words) {
    try {
        const accessToken = await signIn();
        const content = JSON.stringify(words);

        const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/highlightWords.json:/content`;
        const response = await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
                Authorization: `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: content
        });

        if (!response.ok) {
            throw new Error(`Error uploading to OneDrive: ${response.status} ${response.statusText}`);
        }

        console.log("Words successfully saved to OneDrive.");
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
    try {
        let words = await loadWordsFromOneDrive();
        if (!Array.isArray(words)) {
            console.log("No words found in OneDrive, initializing empty array");
            words = [];
        }
        chrome.storage.local.set({ highlightWords: words }, () => {
            console.log("Words synced to local storage.");
            updateWordList(words);
        });
    } catch (error) {
        console.error("Error syncing words:", error);
    }
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
