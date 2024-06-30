import React, { useState, useEffect } from 'react';
import { PublicClientApplication } from '@azure/msal-browser';
import { Client } from '@microsoft/microsoft-graph-client';

const msalConfig = {
    auth: {
        clientId: "ecd171a6-7d57-443d-ba48-6a53d1c8712b",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "https://excel.vxtdemo.com"
    }
};

const msalInstance = new PublicClientApplication(msalConfig);

const App = () => {
    const [isAuthenticated, setIsAuthenticated] = useState(false);
    const [user, setUser] = useState(null);
    const [folders, setFolders] = useState([]);

    useEffect(() => {
        const initializeMsal = async () => {
            try {
                await msalInstance.initialize(); // Ensure MSAL is initialized
                const response = await msalInstance.handleRedirectPromise();
                if (response) {
                    const account = response.account;
                    msalInstance.setActiveAccount(account);
                    setIsAuthenticated(true);
                    setUser(account);
                    fetchOneDriveFolders();
                } else {
                    const currentAccounts = msalInstance.getAllAccounts();
                    if (currentAccounts.length === 1) {
                        msalInstance.setActiveAccount(currentAccounts[0]);
                        setIsAuthenticated(true);
                        setUser(currentAccounts[0]);
                        fetchOneDriveFolders();
                    }
                }
            } catch (error) {
                console.error("V5. Error during MSAL initialization: ", error);
            }
        };
        initializeMsal();
    }, []);

    const login = async () => {
        try {
            const loginRequest = {
                scopes: ["openid", "profile", "User.Read", "Files.Read.All"]
            };
            const response = await msalInstance.loginPopup(loginRequest);
            if (response) {
                const account = response.account;
                msalInstance.setActiveAccount(account);
                setIsAuthenticated(true);
                setUser(account);
                fetchOneDriveFolders();
            }
        } catch (error) {
            console.error("Login error: ", error);
        }
    };

    const logout = () => {
        msalInstance.logout();
        setIsAuthenticated(false);
        setUser(null);
        setFolders([]);
    };

    const fetchOneDriveFolders = async () => {
        try {
            const account = msalInstance.getActiveAccount();
            if (!account) {
                console.error("No active account! Please log in.");
                return;
            }

            const accessTokenRequest = {
                scopes: ["Files.Read.All"],
                account: account
            };
            const accessTokenResponse = await msalInstance.acquireTokenSilent(accessTokenRequest);
            const accessToken = accessTokenResponse.accessToken;

            const client = Client.init({
                authProvider: (done) => {
                    done(null, accessToken);
                }
            });

            const response = await client.api('/me/drive/root/children').get();
            setFolders(response.value);
        } catch (error) {
            console.error("Error fetching OneDrive folders: ", error);
        }
    };

    return (
        <div>
            <h1>MSAL React App v6</h1>
            {!isAuthenticated ? (
                <button onClick={login}>Login with Microsoft</button>
            ) : (
                <div>
                    <h2>Welcome, {user && user.name}!</h2>
                    <button onClick={logout}>Logout</button>
                    <h3>OneDrive Folders:</h3>
                    <ul>
                        {folders.map(folder => (
                            <li key={folder.id}>{folder.name}</li>
                        ))}
                    </ul>
                </div>
            )}
        </div>
    );
};

export default App;
