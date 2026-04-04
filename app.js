// app.js

// Function to manage and display recent apps
function displayRecentApps() {
    const recentAppsContainer = document.getElementById('recent-apps');
    const recentApps = loadAppsFromLocalStorage();

    recentAppsContainer.innerHTML = '';

    recentApps.forEach(app => {
        const appElement = document.createElement('div');
        appElement.className = 'app';
        appElement.innerText = app.name;
        appElement.onclick = () => handleNavigation(app);
        recentAppsContainer.appendChild(appElement);
    });
}

// Function to load apps from localStorage
function loadAppsFromLocalStorage() {
    const apps = localStorage.getItem('recentApps');
    return apps ? JSON.parse(apps) : [];
}

// Function to handle navigation to app detail
function handleNavigation(app) {
    // Logic to navigate to the app's detail page
    window.location.href = `appDetail.html?id=${app.id}`;
}

// Function to update the UI dynamically
function updateUI() {
    displayRecentApps();
    // Add more UI updates as needed.
}

// Execute UI update on load
window.onload = updateUI;