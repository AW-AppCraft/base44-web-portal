// auth.js

// Function to register a new user
function register(username, password) {
    // Registration logic here
    localStorage.setItem(username, JSON.stringify({ password }));
}

// Function to login a user
function login(username, password) {
    const user = JSON.parse(localStorage.getItem(username));
    if (user && user.password === password) {
        localStorage.setItem('currentUser', username);
        return true;
    }
    return false;
}

// Function to logout a user
function logout() {
    localStorage.removeItem('currentUser');
}

// Function to check if a user is logged in
function isLoggedIn() {
    return localStorage.getItem('currentUser') !== null;
}

// Function to get the current logged-in user
function getCurrentUser() {
    return localStorage.getItem('currentUser');
}

// Function to add an app to the recent apps
function addToRecentApps(appName) {
    const recentApps = JSON.parse(localStorage.getItem('recentApps')) || [];
    recentApps.push(appName);
    localStorage.setItem('recentApps', JSON.stringify(recentApps));
}

// Function to subscribe to an app
function subscribeToApp(appName) {
    const subscriptions = JSON.parse(localStorage.getItem('subscriptions')) || [];
    if (!subscriptions.includes(appName)) {
        subscriptions.push(appName);
        localStorage.setItem('subscriptions', JSON.stringify(subscriptions));
    }
}

// Function to unsubscribe from an app
function unsubscribeFromApp(appName) {
    let subscriptions = JSON.parse(localStorage.getItem('subscriptions')) || [];
    subscriptions = subscriptions.filter(app => app !== appName);
    localStorage.setItem('subscriptions', JSON.stringify(subscriptions));
}

// Middleware to require login
function requireLogin() {
    if (!isLoggedIn()) {
        throw new Error('User is not logged in.');
    }
}

// Function to check if the user is admin
function isAdmin() {
    const currentUser = getCurrentUser();
    return currentUser === 'admin'; // Change this logic based on your admin verification
}

export { register, login, logout, isLoggedIn, getCurrentUser, addToRecentApps, subscribeToApp, unsubscribeFromApp, requireLogin, isAdmin };