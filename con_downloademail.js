// A unique namespace for your libraries
var SDK = window.SDK || {};
(function () {
    const varHeaders = {
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Content-Type": "application/json; charset=utf-8",
        "Accept": "application/json",
        "Prefer": "odata.include-annotations=*"
    };
    this.OnClickDownloadEmail = async function (executionContext) {
        console.log("Dowloading email...")

        var emailGuid = executionContext.entityReference.id.replace("{", "").replace("}", "")
        var emailData = await fetch(Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/emails(" + emailGuid + ")", {
            method: "GET",
            headers: varHeaders
        });
        var emailRecord = await emailData.json();
        console.log({ emailRecord });
        var emailFrom = emailRecord.sender;
        var emailTo = emailRecord.torecipients;
        var emailCc = await this.getEmailAddress("cc", executionContext);
        var emailBcc = await this.getEmailAddress("bcc", executionContext);
        var emailSubject = executionContext.getAttribute("subject").getValue();
        var htmlDocument = executionContext.getAttribute("description").getValue();
        //var emlContent = "data:message/rfc822 eml;charset=utf-8,";
        var emlContent = 'From: ' + emailFrom + '\n';
        emlContent += 'To: ' + emailTo + '\n';
        emailCc != "" ? emlContent += 'Cc: ' + emailCc + '\n' : null;
        emailBcc != "" ? emlContent += 'Bcc: ' + emailBcc + '\n' : null;
        emlContent += 'Subject: ' + emailSubject + '\n';
        emlContent += 'X-Unsent: 1' + '\n';
        emlContent += "Content-Type: multipart/mixed; boundary=--boundary_text_string" + "\n\n";
        emlContent += "----boundary_text_string" + "\n";
        emlContent += "Content-Type: text/html; charset=UTF-8" + "\n\n"

        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlDocument, "text/html");
        console.log({ doc });
        emlContent += doc.documentElement.innerHTML

        // Get attachments
        if (emailRecord.attachmentcount > 0) {
            var attachmentData = await fetch(Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/activitymimeattachments?$filter=_objectid_value eq " + emailGuid, {
                method: "GET",
                headers: varHeaders
            });
            var attachmentRecords = await attachmentData.json();
            attachmentRecords.value.forEach(attachment => {
                emlContent += "\n\n----boundary_text_string" + "\n";
                emlContent += "Content-Type: " + attachment.mimetype + "; name=" + attachment.filename + "\n";
                emlContent += "Content-Transfer-Encoding: base64" + "\n"
                emlContent += "Content-Disposition: attachment" + "\n\n";
                emlContent += attachment.body;
            });
            emlContent += "\n\n----boundary_text_string--";
        }
        var text = emlContent;
        var textFile = null;
        var data = new Blob([text], { type: 'text/plain' });
        if (textFile !== null) {
            window.URL.revokeObjectURL(textFile);
        }
        textFile = window.URL.createObjectURL(data);
        var anchor = document.createElement('a');
        var link = document.createTextNode("downloadlink");
        anchor.appendChild(link);
        anchor.href = textFile;
        anchor.id = 'downloadlink';
        // Remove illegal characters from file name
        anchor.download = emailSubject.replace(/[/\\?%*:|"<>]/g, '-') + '.eml';
        anchor.style = "display:none;"; //hidden link
        document.body.appendChild(anchor);
        document.getElementById('downloadlink').click(); //click the link
    }
    this.getEmailAddress = async function (schema, executionContext) {
        var parties = executionContext.getAttribute(schema).getValue();
        var emailAddresses = "";
        for (const party of parties) {
            var partyGuid = party.id.replace("{", "").replace("}", "");
            var userData = await fetch(Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/" + party.entityType + "s(" + partyGuid + ")", {
                method: "GET",
                headers: varHeaders
            });
            var userRecord = await userData.json();
            console.log({ userRecord });
            if (party.entityType == "systemuser") {
                emailAddresses += userRecord.internalemailaddress + ";";
            }
            else {
                emailAddresses += userRecord.emailaddress1 + ";";
            }
        }
        return emailAddresses;
    }
}).call(SDK);