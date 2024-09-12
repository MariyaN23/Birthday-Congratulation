//min date
const minDate = new Date().toISOString().split('T')[0]
const inputMinDate = document.getElementById('minBirthday')
inputMinDate.min = minDate

//max date
const maxDate = new Date().toISOString().split('T')[0]
const inputMaxDate = document.getElementById('maxBirthday')
inputMaxDate.min = maxDate

let stopProcessing = false
let processed = 0
let errorSendingEmail = 0
let all
let alreadySentEmail = 0
let total

const startButton = () => {
    const cancelButton = document.querySelector('.lng-cancel-btn')

    stopProcessing = false
    processed = 0
    errorSendingEmail = 0
    all = 0
    alreadySentEmail = 0
    total = 0

    const progressBar = document.getElementById("progressBar")
    const processedDiv = document.getElementById("processedDiv")
    clearAllClasses(processedDiv)
    progressBar.style.width = 0 + "%"
    progressBar.innerHTML = 0 + "%"

    const minInputValue = new Date(document.getElementById('minBirthday').value)
    const maxInputValue = new Date(document.getElementById('maxBirthday').value)
    const sevenDaysLater = new Date()
    //sevenDaysLater.setDate(sevenDaysLater.getDate() + 7)
    sevenDaysLater.setDate(minInputValue.getDate() + 7)

    const errorDiv = document.getElementById('errorDiv')
    errorDiv.style.display = 'none'
    clearAllClasses(errorDiv)

    const successDiv = document.getElementById('successDiv')
    successDiv.style.display = 'none'
    clearAllClasses(successDiv)


    const isMinInputEmpty = isNaN(minInputValue.getTime())
    const isMaxInputEmpty = isNaN(maxInputValue.getTime())

    /*while (errorDiv.classList.length > 0) {
        errorDiv.classList.remove(errorDiv.classList.item(0));
    }*/
    errorDiv.classList.add("error")

    switch (true) {
        case minInputValue > maxInputValue:
            errorDiv.innerText = 'Min value should not be more than max'
            errorDiv.style.display = 'block'
            errorDiv.classList.add("lng-minMoreThanMax-error")
            break;
        case maxInputValue > sevenDaysLater:
            errorDiv.innerText = 'Max value should not be more than 7 days'
            errorDiv.style.display = 'block'
            errorDiv.classList.add("lng-maxMoreThanSevenDays-error")
            break;
        case isMinInputEmpty && isMaxInputEmpty:
            errorDiv.innerText = 'Select min and max dates'
            errorDiv.style.display = 'block'
            errorDiv.classList.add("lng-selectMinMaxDates-error")
            break;
        case isMinInputEmpty:
            errorDiv.innerText = `Min date shouldn't be empty`
            errorDiv.style.display = 'block'
            errorDiv.classList.add("lng-minDateEmpty-error")
            break;
        case isMaxInputEmpty:
            errorDiv.innerText = `Max date shouldn't be empty`
            errorDiv.style.display = 'block'
            errorDiv.classList.add("lng-maxDateEmpty-error")
            break;
        default:
            authContext.acquireToken(organizationURI, retrieveContacts)
            cancelButton.style.display = 'block'
    }
    changeLanguage(languageId)

    function retrieveContacts(error, token) {
        if (error || !token) {
            console.log('ADAL error occurred: ' + error);
            errorMessage.textContent = 'ADAL error occurred: ' + error;
            return;
        }
        const startDateString = minInputValue.toISOString().slice(0, 10)
        const endDateString = maxInputValue.toISOString().slice(0, 10)
        const filter = `birthdate ge ${startDateString} and birthdate le ${endDateString}`
        const contactsQuery = `/api/data/v8.0/contacts?$select=fullname,emailaddress1,test_last_birthday_congrat,contactid,gendercode&$filter=${filter}`

        const req = new XMLHttpRequest()
        req.open("GET", encodeURI(organizationURI + contactsQuery), true);
        req.setRequestHeader("Authorization", "Bearer " + token);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    const accounts = JSON.parse(this.response).value;
                    success(accounts)
                } else {
                    var error = JSON.parse(this.response).error;
                    console.log(error.message);
                    errorMessage.textContent = error.message;
                }
            }
        }
        req.send();
    }

    var mycontact

    function updateContact(error, token) {
        const contactsQuery = `/api/data/v8.0/contacts(${mycontact.contactid})`
        const payload = JSON.stringify({test_last_birthday_congrat: true})
        const req = new XMLHttpRequest()
        req.open("PATCH", encodeURI(organizationURI + contactsQuery), true);
        //Set Bearer token
        req.setRequestHeader("Authorization", "Bearer " + token);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 204) {
                    console.log('Contact updated successfully')
                } else {
                    var error = JSON.parse(this.response).error;
                    console.log(error.message);
                    errorMessage.textContent = error.message;
                }
            }
        }
        req.send(payload)
    }

    function success(result) {
        total = result.length
        if (!total) {
            successDiv.style.display = 'block'
            successDiv.classList.add("lng-0-contacts")
            cancelButton.style.display = 'none'
            processedDiv.style.display = 'block'
            processedDiv.classList.add("lng-all-contacts-processed")
            changeLanguage(languageId)
            return
        }

        result.forEach(function (contact) {
            if (stopProcessing) {
                return
            }
            mycontact = contact
            const emailAddress = contact.emailaddress1
            const isEmailSentToUser = contact.test_last_birthday_congrat
            const genderCode = contact.gendercode
            const userName = contact.fullname
            let birthdayMessageBody
            let congratulation
            if (languageId === 1033) {
                congratulation = `Happy birthday! I hope all your birthday wishes and dreams come true!`
            }
            if (languageId === 1049) {
                congratulation = `с днем рождения! Надеюсь, все Ваши желания и мечты сбудутся!`
            }
            if (genderCode === 1) {
                birthdayMessageBody = `Sehr geehrter Herr ${userName}, ${congratulation}.`
            } else if (genderCode === 2) {
                birthdayMessageBody = `Sehr geehrte Frau ${userName}, ${congratulation}.`
            }

            function updateEmailStatus() {
                progressBar.style.display = 'block'
                processedDiv.style.display = 'block'

                all = processed + errorSendingEmail + alreadySentEmail
                processedDiv.classList.add("lng-progress")

                const progress = Math.round((all / total) * 100)
                progressBar.style.width = progress + "%"
                progressBar.innerHTML = progress + "%"

                if (all === total) {
                    successDiv.style.display = 'block'
                    successDiv.classList.add("lng-all-contacts-processed")
                    cancelButton.style.display = 'none'
                    clearAllClasses(processedDiv)
                    processedDiv.classList.add("lng-processed-contacts")
                }
                changeLanguage(languageId)
            }

            if (!isEmailSentToUser) {
                Email.send({
                    Host: "smtp.elasticemail.com",
                    Username: "sellaite505@gmail.com",
                    Password: "780A6E10DC3187E084E5BAFDCB85B66F2AAF",
                    To: emailAddress,
                    From: "sellaite505@gmail.com",
                    Subject: "Happy Birthday!",
                    Body: birthdayMessageBody
                }).then(
                    message => {
                        if (message === 'OK') {
                            processed++
                            authContext.acquireToken(organizationURI, updateContact)
                        } else {
                            errorSendingEmail++
                        }
                        updateEmailStatus()
                    }
                )
            } else {
                alreadySentEmail++
                updateEmailStatus()
            }
            changeLanguage(languageId)
        })
    }
}

function clearAllClasses(processedDiv) {
    while (processedDiv.classList.length > 0) {
        processedDiv.classList.remove(processedDiv.classList.item(0))
    }
}

const cancelProcessing = () => {
    const progressBar = document.getElementById("progressBar")
    progressBar.style.display = 'none'
    if (confirm(`Are you sure you want to cancel the processing?`)) {
        stopProcessing = true
        const processedDiv = document.getElementById("processedDiv")
        clearAllClasses(processedDiv)
        processedDiv.style.display = 'block'
        processedDiv.classList.add("lng-cancel-processing")
        const successDiv = document.getElementById("successDiv")
        clearAllClasses(successDiv)
        successDiv.style.display = 'block'
        successDiv.classList.add("lng-all-cancelled")
        changeLanguage(languageId)
    }
}