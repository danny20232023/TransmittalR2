function navigateToEmail() {
    var transmittalId = Xrm.Page.data.entity.getId().replace(/[{}]/g, "");
    var transmittalNumber = Xrm.Page.data.entity.getPrimaryAttributeValue();
    var transmittalTitleAttr = Xrm.Page.data.entity.attributes.get("new_title");
    var transmittalProjectNumberAttr = Xrm.Page.data.entity.attributes.get("crbd5_projectnumber");

    var transmittalTitleValue = transmittalTitleAttr ? transmittalTitleAttr.getValue() : ""; 
    var transmittalProjectNumberValue = transmittalProjectNumberAttr ? transmittalProjectNumberAttr.getValue() : ""; 

    console.log("Transmittal ID:", transmittalId);
    console.log("Transmittal Number:", transmittalNumber);
    console.log("Transmittal Title:", transmittalTitleValue);
    console.log("Transmittal Project Number:", transmittalProjectNumberValue);

    // Retrieve all CC emails
    Xrm.WebApi.retrieveMultipleRecords("new_projecttransmittalcc", "?$select=new_email&$filter=new_iscc eq true").then(
        function (ccResult) {
            var emailList = [];

            for (var i = 0; i < ccResult.entities.length; i++) {
                var email = ccResult.entities[i]["new_email"];
                if (email) {
                    emailList.push(email);
                }
            }

            var transmittalCC = emailList.join(", ");
            console.log("Transmittal CC:", transmittalCC);

          // Retrieve related Project Transmittal SendTo table column name:emails, via many-to-many relationship to Transmittal Table
Xrm.WebApi.retrieveRecord("crbd5_transmittalregister", transmittalId,
    "?$expand=new_TransmittalSendtoUID($select=new_emailaddress)").then(
    function (sendToResult) {
        var sendToEmails = [];

        // Must match the $expand property name
        if (sendToResult.new_TransmittalSendtoUID) {
            sendToResult.new_TransmittalSendtoUID.forEach(function (record) {
                if (record.new_emailaddress) {
                    sendToEmails.push(record.new_emailaddress);
                }
            });
        }

        var sendToValue = sendToEmails.join(", ");
        console.log("Transmittal SendTo:", sendToValue);

                    // Create the new record with all populated fields
                    var entityFormOptions = {
                        entityName: "new_projecttransmittalsentemail"
                    };

                    var formParameters = {
                        new_emailtransmittaluid: transmittalId,
                        new_mail: transmittalTitleValue,
                        new_cc: transmittalCC,
                        new_sendto: sendToValue,
                        new_subject: "Transmittal " + transmittalNumber + " for " + transmittalTitleValue,
                        new_body:
                            "<div style='font-family: Aptos, sans-serif; font-size: 15px;'>" +
                            "<p style='margin-bottom: 10px;'>Hi Client<strong style='color: blue;'></strong>,</p>" +
                            "<p style='margin-bottom: 5px;'>Please find attached transmitted documents for <strong style='color: green;'>" + transmittalTitleValue + "</strong>.</p>" +
                            "<p style='margin-bottom: 10px;'>Can you please acknowledge receipt and return to <a href='mailto:documentcontrol@papsolutions.com.au'>documentcontrol@papsolutions.com.au</a>.</p>" +
                            "<p style='margin-bottom: 10px;'>Kind Regards,</p>" +
                            "<p><strong>Pap Solutions Pty Ltd Document Control</strong></p>" +
                            "</div>"
                    };

                    console.log("Opening new Sent Email Form with parameters:", formParameters);
                    Xrm.Navigation.openForm(entityFormOptions, formParameters);
                },
                function (error) {
                    console.error("Error retrieving SendTo emails:", error.message);
                }
            );
        },
        function (error) {
            console.error("Error retrieving Project Transmittal CC emails:", error.message);
        }
    );
}

// Expose to global scope
window.navigateToEmail = navigateToEmail;
