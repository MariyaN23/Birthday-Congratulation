//Set these variables to match your environment
var organizationURI = "https://orgd9f95318.crm11.dynamics.com"; //The URL to connect to CRM Online
var tenant = "qsolutions349.onmicrosoft.com"; //The name of the Azure AD organization you use
var clientId = "d5cd66c0-93d6-4051-acc8-6b0f95bbb09f"; //The ClientId you got when you registered the application
var pageUrl = "https://orgd9f95318.crm11.dynamics.com"; //The URL of this page in your development environment when debugging.
var user, authContext, message, errorMessage, loginButton, logoutButton, accountsTable,
    accountsTableBody;
var userId
var languageId
//Configuration data for AuthenticationContext
var endpoints = {
    orgUri: organizationURI
};
window.config = {
    tenant: tenant,
    clientId: clientId,
    postLogoutRedirectUri: pageUrl,
    endpoints: endpoints,
    cacheLocation: 'localStorage', // enable this for IE, as sessionStorage does not work for localhost.
};
document.onreadystatechange = function () {
    if (document.readyState === "complete") {
        //Set DOM elements referenced by scripts
        message = document.getElementById("message");
        errorMessage = document.getElementById("errorMessage");
        loginButton = document.getElementById("login");
        logoutButton = document.getElementById("logout");
        accountsTable = document.getElementById("accountsTable");
        accountsTableBody = document.getElementById("accountsTableBody");
        //Event handlers on DOM elements
        loginButton.addEventListener("click", login);
        logoutButton.addEventListener("click", logout);
        authenticate();
        if (user) {
            loginButton.style.display = "none";
            logoutButton.style.display = "block";
            var helloMessage = document.createElement("p");
            helloMessage.textContent = "Hello, " + user.profile.name;
            helloMessage.classList.add("lng-greetingUser")
            message.appendChild(helloMessage)

            const div = document.getElementById('openBirthdayCongratulation')
            div.style.display = 'block'
            authContext.acquireToken(organizationURI, getUsersId)
        } else {
            loginButton.style.display = "block";
            logoutButton.style.display = "none";
        }
    }
}

// Function that manages authentication
function authenticate() {
    //OAuth context
    authContext = new AuthenticationContext(config);
    // Check For & Handle Redirect From AAD After Login
    var isCallback = authContext.isCallback(window.location.hash);
    if (isCallback) {
        authContext.handleWindowCallback();
    }
    var loginError = authContext.getLoginError();
    if (isCallback && !loginError) {
        window.location = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST);
    } else {
        errorMessage.textContent = loginError;
    }
    user = authContext.getCachedUser();
}

//function that logs in the user
function login() {
    authContext.login()
    //authContext.acquireToken(organizationURI, getUsersLanguage)
}

//function that logs out the user
function logout() {
    authContext.logOut();
    accountsTable.style.display = "none";
    accountsTableBody.innerHTML = "";
}

//language settings
function getUsersId(error, token) {
    if (error || !token) {
        console.log('ADAL error occurred: ' + error);
        errorMessage.textContent = 'ADAL error occurred: ' + error;
        return;
    }
    const params = `/api/data/v8.0/WhoAmI`
    const req = new XMLHttpRequest()
    req.open("GET", encodeURI(organizationURI + params), true);
    req.setRequestHeader("Authorization", "Bearer " + token);
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status === 200) {
                userId = JSON.parse(this.response).UserId
                authContext.acquireToken(organizationURI, getUsersLanguageSettings)
            } else {
                var error = JSON.parse(this.response).error;
                console.log(error.message);
                errorMessage.textContent = error.message;
            }
        }
    }
    req.send()
}

function getUsersLanguageSettings(error, token) {
    if (error || !token) {
        console.log('ADAL error occurred: ' + error);
        errorMessage.textContent = 'ADAL error occurred: ' + error;
        return;
    }
    const params = `/api/data/v9.0/usersettingscollection(${userId})`
    const req = new XMLHttpRequest()
    req.open("GET", encodeURI(organizationURI + params), true);
    //Set Bearer token
    req.setRequestHeader("Authorization", "Bearer " + token);
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status === 200) {
                languageId = JSON.parse(this.response).uilanguageid;
                changeLanguage(languageId)
            } else {
                var error = JSON.parse(this.response).error;
                console.log(error.message);
                errorMessage.textContent = error.message;
            }
        }
    }
    req.send()
}

//language
const languages = {
    "title": {
        "en": "Birthday congratulation",
        "ru": "Поздравление с днем рождения"
    },
    "login-btn": {
        "en": "Login",
        "ru": "Войти"
    },
    "logout-btn": {
        "en": "Logout",
        "ru": "Выйти"
    },
    "greetingUser": {
        "en": "Hello, ",
        "ru": "Здравствуйте, "
    },
    "start-btn": {
        "en": "Start",
        "ru": "Начать"
    },
    "cancel-btn": {
        "en": "Cancel",
        "ru": "Отмена"
    },
    "min-date": {
        "en": "Start date:",
        "ru": "Начальная дата:"
    },
    "max-date": {
        "en": "End date:",
        "ru": "Конечная дата:"
    },

    //error messages
    "minMoreThanMax-error": {
        "en": "Min value should not be more than max",
        "ru": "Минимальное значение не должно быть больше максимального"
    },
    "maxMoreThanSevenDays-error": {
        "en": "Max value should not be more than 7 days",
        "ru": "Максимальное значение не должно превышать 7 дней"
    },
    "selectMinMaxDates-error": {
        "en": "Select min and max dates",
        "ru": "Выберите минимальную и максимальную даты"
    },
    "minDateEmpty-error": {
        "en": "Min date shouldn't be empty",
        "ru": "Минимальная дата не должна быть пустой"
    },
    "maxDateEmpty-error": {
        "en": "Max date shouldn't be empty",
        "ru": "Максимальная дата не должна быть пустой"
    },
    //processed
    "processed-contacts": {
        "en": (processed, total, errorSendingEmail, alreadySentEmail) => `Processed contacts ${processed} from ${total}, error ${errorSendingEmail}, already sent ${alreadySentEmail} emails`,
        "ru": (processed, total, errorSendingEmail, alreadySentEmail) => `Обработано контактов ${processed} из ${total}, ${errorSendingEmail} ошибок, ранее отправленных сообщений: ${alreadySentEmail}`
    },
    "all-contacts-processed": {
        "en": "All contacts processed!",
        "ru": "Все контакты обработаны!"
    },
    "progress": {
        "en": (processed, total) => `Processed ${processed} contacts from ${total}`,
        "ru": (processed, total) => `Обработано ${processed} контактов из ${total}`
    },
    "0-contacts": {
        "en": "0 contacts found with this dates!",
        "ru": "Найдено 0 контактов с этими датами!"
    },
    //cancel
    "cancel-processing": {
        "en": (processed, errorSendingEmail) => `Processing is cancelled successfully. ${processed} contacts are processed,
        ${processed} congratulation e-mails were successfully sent (sending of ${errorSendingEmail} e-mails failed).`,
        "ru": (processed, errorSendingEmail) => `Обработка успешно отменена. ${processed} контактов обработано, 
        ${processed} писем успешно отправлено (отправка ${errorSendingEmail} писем не удалась).`
    },
    "all-cancelled": {
        "en": "Processing cancelled!",
        "ru": "Обработка отменена!"
    },
}

function changeLanguage(languageCode) {
    let code
    if (languageCode === 1033) {
        code = "en"
    }
    if (languageCode === 1049) {
        code = "ru"
    }
    document.querySelector("title").innerHTML = languages["title"][code]
    for (let key in languages) {
        let element = document.querySelector(`.lng-${key}`)
        if (element) {
            switch (key) {
                case "greetingUser":
                    element.innerHTML = `${languages[key][code]}${user.profile.name}`
                    break;
                case "processed-contacts":
                    element.innerHTML = languages[key][code](processed, total, errorSendingEmail, alreadySentEmail)
                    break;
                case "progress":
                    element.innerHTML = languages[key][code](processed, total)
                    break;
                case "cancel-processing":
                    element.innerHTML = languages[key][code](processed, errorSendingEmail)
                    break;
                default:
                    element.innerHTML = languages[key][code]
            }
        }
    }
}