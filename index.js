const excel = require("exceljs");
const nodemailer = require("nodemailer");

class MailingApp {
    
}

const makeRecipients = (e) => {
    let workbook = new excel.Workbook();
    let column = document.getElementById("column").value.toUpperCase();
    let sheet = Number.parseInt(document.getElementById('sheet').value);
    let excelFile = document.getElementById("excelFile").files[0];
    const emailsList = document.getElementById("emailsList");
    const emailsPattern = /\S+@\S+.\S+/;

    workbook.xlsx.readFile(excelFile.path).then(function () {
        emailsList.value = 'Список получателей:\n';
        var worksheet = workbook.getWorksheet(sheet);
        for (let i = 1; i <= worksheet.actualRowCount; i++) {
            let cell = worksheet.getCell(column + i);
            if (cell.value && cell.value !== null) {
                let cellValue = worksheet.getCell(column + i).value;

                /* if it returns an list of objects, take emails by key, split it if returns several emails  */
                if (cellValue.richText) {
                    cellValue.richText.forEach(el => {
                        el.text.includes('\n') ?
                            el.text.split('\n').forEach(elem => recipients.push(elem)) :
                            recipients.push(el.text);
                    });

                    /* if returns object, take emails by key, split it if returns several emails */
                } else if (cellValue.text) {
                    cellValue.text.includes('\n') ?
                        cellValue.text.split('\n').forEach(elem => recipients.push(elem)) :
                        recipients.push(cellValue.text)

                    /* if string has several emails, split it */
                } else {
                    cellValue.includes('\n') ?
                        cellValue.split('\n').forEach(elem => recipients.push(elem)) :
                        recipients.push(cellValue);
                }
            }
        }
        /* take only emails and show them in the app */
        recipients = recipients.filter(c => emailsPattern.test(c));
        recipients.forEach(el => {
            emailsList.value += `${el}\n`
        });
    }).catch((e) => console.log(e));
    return recipients;
}

function sendMailsHandler(elem, recipients) {
    elem.addEventListener('click', (e) => {
        const from = document.getElementById("from").value;
        const password = document.getElementById("password").value;
        const subject = document.getElementById("subject").value;
        const sender = document.getElementById("sender").value;
        const text = document.getElementById("text").value;
        const files = document.getElementById("files").files;
        const emailsList = document.getElementById("emailsList");
        emailsList.value = '';
        var filesList = [];

        for (let i = 0; i < files.length; i++) {
            let file = {
                filename: files[i].name,
                path: files[i].path
            };
            filesList.push(file)
        }
        let transport = nodemailer.createTransport({
            host: 'mail.nic.ru',
            auth: {
                user: from,
                pass: password
            }
        });
        recipients.forEach(recipient => {
            let info = transport.sendMail({
                from: `${sender} <${from}>`,
                to: recipient,
                subject: subject,
                text: text,
                attachments: filesList
            });
            emailsList.value += `message sent: ${recipient}\n`;
        })
    })
}

function changeFileInputHandler() {
    var inputs = document.querySelectorAll('.inputs');
    inputs.forEach((input) => { 
        var label = input.nextElementSibling;
        var labelVal = label.innerHTML;

        input.addEventListener('change', function (e) {
            var fileName = '';
            input.files && input.files.length > 1 ?
                fileName = `${input.files.length} files selected` :
                fileName = input.files[0].name;

            fileName ? label.innerHTML = fileName : label.innerHTML = labelVal;
        });
    });
}

var recipients = [];
const makeRecipientsButton = document.getElementById("makeRecipientsButton");
const sendEmailsButton = document.getElementById("sendEmailsButton");

makeRecipientsButton.addEventListener("click", makeRecipients);
sendMailsHandler(sendEmailsButton, recipients);
changeFileInputHandler();
