const msalConfig = {
    auth: {
        clientId: "2a0d49d5-3f9a-4c80-8e21-4670bf6f53b3",
        authority: "https://login.microsoftonline.com/exaring.de",
        redirectUri: "https://exaring.github.io/logos/signature.html",
    },
};

const loginRequest = { scopes: ["User.Read"] };
const msalInstance = new msal.PublicClientApplication(msalConfig);

/**
 * Sign in with Office 365 and fetch user data.
 */
function signIn() {
    msalInstance.loginPopup(loginRequest)
        .then(response => {
            const account = msalInstance.getAllAccounts()[0];
            if (account) {
                msalInstance.setActiveAccount(account);
                console.log("Signed in:", account);
                getUserData();
            }
        })
        .catch(error => {
            console.error("Login error:", error);
        });
}

function getUserData() {
    msalInstance.acquireTokenSilent(loginRequest)
        .then(tokenResponse => {
            fetch("https://graph.microsoft.com/v1.0/me?$select=givenName,surname,mail,jobTitle,department,mobilePhone", {
                headers: { Authorization: `Bearer ${tokenResponse.accessToken}` },
            })
                .then(response => response.json())
                .then(data => populateForm(data))
                .catch(error => console.error("Graph API error:", error));
        })
        .catch(error => console.error("Token acquisition error:", error));
}

/**
 * Populate the form with user data.
 * @param {Object} data - User data.
 */
function populateForm(data) {
    document.getElementById("firstname").value = data.givenName || "";
    document.getElementById("lastname").value = data.surname || "";
    document.getElementById("email").value = data.mail || "";
    document.getElementById("title").value = data.jobTitle || "";
    document.getElementById("department").value = data.department || "";
    document.getElementById("phone").value = data.mobilePhone || "";
}

function generateSignature() {
    const firstname = document.getElementById('firstname').value.trim();
    const lastname = document.getElementById('lastname').value.trim();
    const email = document.getElementById('email').value.trim();
    const title = document.getElementById('title').value.trim();
    const department = document.getElementById('department').value.trim();
    const phone = document.getElementById('phone').value.trim();
    const language = document.getElementById('language').value.trim();

    if (!firstname || !lastname || !email) {
        alert('Please fill out the mandatory fields (First Name, Last Name, Email).');
        return;
    }

    let titleDepartment = '';
    if (title || department) {
        titleDepartment = `<span style=\"color: gray;\">${title}${title && department ? ' - ' : ''}${department}</span>`;
    }

    let phoneLink = '';
    if (phone) {
        phoneLink = `<a href=\"tel:${phone.replace(/[^\d+]/g, '')}\" style=\"color: #000000; text-decoration: none !important;\">${phone}</a><br>`;
    }

    const germanFooter = `
                <div style=\"margin: 0; font-size: 10px; color: gray;\">
                    EXARING AG, Leopoldstraße 236, 80807 München<br>
                    Vorstand: Christoph Bellmer, Markus Haertenstein, Robert Laier | Amtsgericht München, HRB 205601<br>
                    Vorsitzender des Aufsichtsrats: Christoph Vilanek
                </div>
            `;

    const englishFooter = `
                <div style=\"margin: 0; font-size: 10px; color: gray;\">
                    EXARING AG, Leopoldstrasse 236, 80807 Munich, Germany<br>
                    Executive Board: Christoph Bellmer, Markus Haertenstein, Robert Laier | District Court of Munich, HRB 205601<br>
                    Chairman of the Supervisory Board: Christoph Vilanek
                </div>
            `;

    const footer = language === 'de' ? germanFooter : englishFooter;


    const signatureHTML = `
                <div style=\"font-family: Arial, sans-serif; font-size: 12px; line-height: 1.2;\">
                    <div style=\"margin: 0;\">
                        --<br>
                        <span style=\"font-size: 14px; font-weight: bold;\">${firstname} ${lastname}</span> ${titleDepartment ? '| ' + titleDepartment : ''}<br>
                        <a href=\"mailto:${email}\" style=\"color: #000000; text-decoration: none;\">${email}</a><br>
                        ${phoneLink}
                    </div>
                    <div style=\"margin: 5px 0;\">
                        <img src=\"/waipu-tv-logo-dunkel@3x.png\"
                             alt=\"waipu.tv Logo\" style=\"height: 26px; display: inline-block;\">
                    </div>
                    ${footer}
                    <div style=\"margin: 5px 0; font-size: 10px;\">
                        <a href=\"https://www.exaring.de\" style=\"color: #2789FC; text-decoration: none;\">www.exaring.de</a> |
                        <a href=\"https://www.waipu.tv\" style=\"color: #2789FC; text-decoration: none;\">www.waipu.tv</a>
                    </div>
                </div>
            `;
    document.getElementById('signature-output').innerHTML = signatureHTML;
    document.getElementById('signature-container').style.display = 'block';
    document.getElementById('html-output').innerText = signatureHTML;
    document.getElementById('html-container').style.display = 'block';

    const textOnlySignature = document.getElementById('signature-output').innerText;
    document.getElementById('text-only-output').innerText = textOnlySignature;
    document.getElementById('text-container').style.display = 'block';
}

/**
 * Copy the content of an element to the clipboard.
 * @param {string} elementId - The ID of the element to copy.
 */
function copyToClipboard(elementId) {
    const content = document.getElementById(elementId).innerText;
    navigator.clipboard.writeText(content)
        .then(() => alert('Content copied to clipboard!'))
        .catch(err => alert('Failed to copy content: ' + err));
}

/**
 * Copy rich HTML content to the clipboard.
 * @param {string} elementId - The ID of the element to copy.
 */
function copyToClipboardRich(elementId) {
    const element = document.getElementById(elementId);
    const html = element.outerHTML;

    const clipboardItem = new ClipboardItem({
        'text/html': new Blob([html], { type: 'text/html' }),
        'text/plain': new Blob([element.innerText], { type: 'text/plain' })
    });

    navigator.clipboard.write([clipboardItem])
        .then(() => alert('Content copied to clipboard!'))
        .catch(error => alert('Failed to copy content: ' + error));
}




