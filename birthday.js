//min date
const minDate = new Date().toISOString().split('T')[0]
const inputMinDate = document.getElementById('minBirthday')
inputMinDate.min = minDate

//max date
const maxDate = new Date().toISOString().split('T')[0]
const inputMaxDate = document.getElementById('maxBirthday')
inputMaxDate.min = maxDate

var Email = {
    send: function (a) {
        return new Promise(function (n, e) {
            a.nocache = Math.floor(1e6 * Math.random() + 1), a.Action = "Send";
            var t = JSON.stringify(a);
            Email.ajaxPost("https://smtpjs.com/v3/smtpjs.aspx?", t, function (e) {
                n(e)
            })
        })
    }, ajaxPost: function (e, n, t) {
        var a = Email.createCORSRequest("POST", e);
        a.setRequestHeader("Content-type", "application/x-www-form-urlencoded"), a.onload = function () {
            var e = a.responseText;
            null != t && t(e)
        }, a.send(n)
    }, ajax: function (e, n) {
        var t = Email.createCORSRequest("GET", e);
        t.onload = function () {
            var e = t.responseText;
            null != n && n(e)
        }, t.send()
    }, createCORSRequest: function (e, n) {
        var t = new XMLHttpRequest;
        return "withCredentials" in t ? t.open(e, n, !0) : "undefined" != typeof XDomainRequest ? (t = new XDomainRequest).open(e, n) : t = null, t
    }
}

let stopProcessing = false
let processed = 0
let errorSendingEmail = 0
let all
let alreadySentEmail = 0

const startButton = () => {
    const cancelButton = document.querySelector('.lng-cancel-btn')

    stopProcessing = false
    processed = 0
    errorSendingEmail = 0
    all = 0
    alreadySentEmail = 0

    const progressBar = document.getElementById("progressBar")
    const processedDiv = document.getElementById("processedDiv")
    progressBar.style.width = 0 + "%"
    progressBar.innerHTML = 0 + "%"

    const minInputValue = new Date(document.getElementById('minBirthday').value)
    const maxInputValue = new Date(document.getElementById('maxBirthday').value)
    const sevenDaysLater = new Date()
    sevenDaysLater.setDate(sevenDaysLater.getDate() + 7)

    const errorDiv = document.getElementById('errorDiv')
    errorDiv.style.display = 'none'

    const successDiv = document.getElementById('successDiv')
    successDiv.style.display = 'none'

    const isMinInputEmpty = isNaN(minInputValue.getTime())
    const isMaxInputEmpty = isNaN(maxInputValue.getTime())

    while (errorDiv.classList.length > 0) {
        errorDiv.classList.remove(errorDiv.classList.item(0));
    }
    errorDiv.classList.add("error")

    if (minInputValue > maxInputValue) {
        errorDiv.innerText = 'Min value should not be more than max'
        errorDiv.style.display = 'block'
        errorDiv.classList.add("lng-minMoreThanMax-error")
    } else if (maxInputValue > sevenDaysLater) {
        errorDiv.innerText = 'Max value should not be more than 7 days'
        errorDiv.style.display = 'block'
        errorDiv.classList.add("lng-maxMoreThanSevenDays-error")
    } else if (isMinInputEmpty && isMaxInputEmpty) {
        errorDiv.innerText = 'Select min and max dates'
        errorDiv.style.display = 'block'
        errorDiv.classList.add("lng-selectMinMaxDates-error")
    } else if (isMinInputEmpty) {
        errorDiv.innerText = `Min date shouldn't be empty`
        errorDiv.style.display = 'block'
        errorDiv.classList.add("lng-minDateEmpty-error")
    } else if (isMaxInputEmpty) {
        errorDiv.innerText = `Max date shouldn't be empty`
        errorDiv.style.display = 'block'
        errorDiv.classList.add("lng-maxDateEmpty-error")
    } else {
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
                    console.log(accounts)
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
        const total = result.length

        progressBar.style.display = 'block'
        processedDiv.style.display = 'block'
        processedDiv.innerText = `Processed contacts ${processed} from ${total}`

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

            if (genderCode === 1) {
                birthdayMessageBody = `Sehr geehrter Herr ${userName},
                            Ich möchte Ihnen herzlich gratulieren zu Ihrem jüngsten Erfolg.
                            Mögen weiterhin viele positive Momente und Errungenschaften Ihren Weg kreuzen.`
            } else if (genderCode === 2) {
                birthdayMessageBody = `Sehr geehrte Frau ${userName},
                            Ich möchte Ihnen herzlich zu Ihren jüngsten Erfolgen gratulieren.
                            Mögen weiterhin viele positive Momente und Errungenschaften Ihren Weg begleiten.`
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
                        all = processed + errorSendingEmail + alreadySentEmail
                        processedDiv.innerText = `Processed contacts ${processed} from ${total}, error ${errorSendingEmail}, already sent ${alreadySentEmail}`
                        const progress = Math.round((all / total) * 100)
                        progressBar.style.width = progress + "%"
                        progressBar.innerHTML = progress + "%"

                        if (all === total) {
                            processedDiv.innerText = `${total} contacts are processed, ${processed} congratulation
                                    e-mails were successfully sent (sending of ${errorSendingEmail} e-mails failed, already sent emails ${alreadySentEmail}).`
                            successDiv.innerText = 'All emails sent!'
                            successDiv.style.display = 'block'
                        }
                    }
                )
            } else {
                alreadySentEmail++
                all = processed + errorSendingEmail + alreadySentEmail
                processedDiv.innerText = `Processed contacts ${processed} from ${total}, error ${errorSendingEmail}, already sent ${alreadySentEmail}`
                const progress = Math.round((all / total) * 100)
                progressBar.style.width = progress + "%"
                progressBar.innerHTML = progress + "%"

                if (all === total) {
                    processedDiv.innerText = `${total} contacts are processed, ${processed} congratulation
                                    e-mails were successfully sent (sending of ${errorSendingEmail} e-mails failed, already sent emails ${alreadySentEmail}).`
                    successDiv.innerText = 'All emails sent!'
                    successDiv.style.display = 'block'
                }
            }
        })
    }
}

const cancelProcessing = () => {
    const progressBar = document.getElementById("progressBar")
    progressBar.style.display = 'none'
    if (confirm("Are you sure you want to cancel the processing?")) {
        stopProcessing = true
        const processedDiv = document.getElementById("processedDiv")
        processedDiv.innerText = `Processing is cancelled successfully. ${processed} contacts are processed,
        ${processed} congratulation e-mails were successfully sent (sending of ${errorSendingEmail} e-mails failed).`
        const successDiv = document.getElementById("successDiv")
        successDiv.innerText = 'Processing Cancelled'
        successDiv.style.display = 'block'
    }
}