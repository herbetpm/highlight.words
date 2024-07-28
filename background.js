
chrome.runtime.onInstalled.addListener(() => {
    console.log("Extension installed");
});

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "saveWordsToOneDrive") {
        saveWordsToOneDrive(request.words);
    }
});

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

async function signIn() {
    const msalConfig = {
        auth: {
            clientId: "e43923d1-c400-4588-843b-3cb5eb99107a",
            authority: "https://login.microsoftonline.com/common",
            redirectUri: "https://herbetpm.github.io/highlight.words/options.html"
        }
    };

    const msalInstance = new msal.PublicClientApplication(msalConfig);

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
