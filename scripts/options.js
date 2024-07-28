
const msalConfig = {
    auth: {
        clientId: "e43923d1-c400-4588-843b-3cb5eb99107a",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "https://herbetpm.github.io/highlight.words/index.html"  // O `options.html` si lo prefieres
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
    await handleRedirect();
    // Cargar y actualizar lista de palabras aquí, si es necesario
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
