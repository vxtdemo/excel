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
                console.error("VX. Error during MSAL initialization: ", error);
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


    const fetchAllItems = async (client, folderId) => {
        let items = [];
        let response = await client.api(`/me/drive/items/${folderId}/children`).get();
        items = items.concat(response.value);
    
        while (response['@odata.nextLink']) {
            response = await client.api(response['@odata.nextLink']).get();
            items = items.concat(response.value);
        }
    
        return items;
    };
    
    const fetchFolderPermissions = async (client, folderId) => {
        try {
            const permissionsResponse = await client.api(`/me/drive/items/${folderId}/permissions`).get();
            return permissionsResponse.value;
        } catch (error) {
            console.error("Error fetching folder permissions: ", error);
            return [];
        }
    };
    
    const getTotalFolderSize = async (client, folderId) => {
        try {
            const items = await fetchAllItems(client, folderId);
            let totalSize = 0;
    
            for (const item of items) {
                if (item.folder) {
                    totalSize += await getTotalFolderSize(client, item.id);
                } else {
                    totalSize += item.size;
                }
            }
    
            if (items.length > 200) {
                console.warn(`Folder with ID ${folderId} contains more than 200 items. Total size may not be accurate.`);
            }
    
            return totalSize;
        } catch (error) {
            console.error("Error calculating folder size: ", error);
            return 0;
        }
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

            // const response = await client.api('/me/drive/root/children').get();
            // setFolders(response.value);

            // Fetch both owned and shared items
            const driveResponse = await client.api('/me/drive/root/children').get();
            const sharedResponse = await client.api('/me/drive/sharedWithMe').get();

            const combinedFolders = [...driveResponse.value, ...sharedResponse.value];
        
            // Fetch size and type of folders
            for (const folder of combinedFolders) {
                folder.type = folder.remoteItem ? 'Shared' : 'Owned';
                folder.size = await getTotalFolderSize(client, folder.id);
                folder.permissions = await fetchFolderPermissions(client, folder.id);
            }
    
            combinedFolders.sort((a, b) => a.name.localeCompare(b.name));
            setFolders(combinedFolders);
        } catch (error) {
            console.error("Error fetching OneDrive folders: ", error);
        }
    };

    return (
        <div>
            <h1>MSAL React App v12 w permissions (in GHP)</h1>
            {!isAuthenticated ? (
                <button onClick={login}>Login with Microsoft</button>
            ) : (
                <div>
                    <h2>Welcome, {user && user.name}!</h2>
                    <button onClick={logout}>Logout</button>
                    <h3>OneDrive Folders:</h3>
                    <ul>
                        {folders.map(folder => (
                            <li key={folder.id}>
                                {folder.name} - {folder.size ? `Size: ${folder.size} bytes` : 'No size data'} - {folder.type}
                                <ul>
                                    {folder.permissions && folder.permissions.length > 0 ? (
                                        folder.permissions.map(permission => (
                                            <li key={permission.id}>
                                                {permission.grantedTo ? permission.grantedTo.user.displayName : 'Unknown'} - {permission.roles.join(", ")}
                                            </li>
                                        ))
                                    ) : (
                                        <li>No permissions data</li>
                                    )}
                                </ul>
                            </li>
                        ))}
                    </ul>
                </div>
            )}
        </div>
    );
};

export default App;
