function sendEmailWithDelay() {
    try {
        // Ensure Xrm.Page is available
        if (!Xrm || !Xrm.Page || !Xrm.Page.data || !Xrm.Page.data.entity) {
            throw new Error("Xrm.Page.data.entity is unavailable.");
        }

        // Get the form context from Xrm.Page
        var formContext = Xrm.Page;

        // Get the record ID and entity logical name (of the current record)
        var entityId = formContext.data.entity.getId().replace("{", "").replace("}", ""); // Remove braces for Web API
        var entityLogicalName = formContext.data.entity.getEntityName();

        // --- Email Details (Customize these) ---
        var emailSubject = "Document Transmittal from Pap Solutions for " + formContext.data.entity.getPrimaryAttributeValue(); // Example: Uses primary field for subject
        var emailBody = "Dear Recipient,<br/><br/>Please find attached the document transmittal related to your project. <br/><br/>Regards,<br/>Pap Solutions";

        // Define email recipients (example: sending to current user, modify as needed)
        // You'll need to fetch the recipient's GUID and logical name.
        // For demonstration, let's assume you want to send to a contact record related to the current record.
        // You would typically retrieve this from a lookup field on your form.
        
        // Example: If you have a lookup field to a 'contact' on your form named 'new_recipientcontact'
        // var recipientContactLookup = formContext.getAttribute("new_recipientcontact").getValue();
        // var recipientContactId = recipientContactLookup ? recipientContactLookup[0].id.replace("{", "").replace("}", "") : null;
        // var recipientContactLogicalName = recipientContactLookup ? recipientContactLookup[0].entityType : null;

        // For simplicity, let's assume we're sending to the current user (SystemUser).
        // You'll likely want to send to specific contacts or users.
        // You can get the current user's ID: Xrm.Utility.getGlobalContext().userSettings.userId
        var currentUserId = Xrm.Utility.getGlobalContext().userSettings.userId.replace("{", "").replace("}", "");

        var emailActivityParties = [
            {
                // From (Current User)
                "partyid_systemuser@odata.bind": "/systemusers(" + currentUserId + ")",
                "participationtypemask": 1 // From
            }
            // Add To, CC, BCC parties as needed
            /*
            , {
                // To (Example: a specific Contact)
                "partyid_contact@odata.bind": "/contacts(" + recipientContactId + ")",
                "participationtypemask": 2 // To
            }
            */
        ];

        // Prepare email record data
        var emailData = {
            "subject": emailSubject,
            "description": emailBody,
            "regardingobjectid_" + entityLogicalName + "@odata.bind": "/" + entityLogicalName + "s(" + entityId + ")", // Associate email with current record
            "email_activity_parties": emailActivityParties,
            "directioncode": true // Outgoing email
        };

        // 1. Create the Email Activity Record
        Xrm.WebApi.createRecord("email", emailData).then(
            function success(emailResult) {
                var emailId = emailResult.id;
                console.log("Email record created with ID: " + emailId);

                // 2. Send the Email Activity
                var sendEmailRequest = {
                    "IssueSend": true // Set to true to send the email immediately
                };

                Xrm.WebApi.online.execute(new Xrm.WebApi.online.boundAction("SendEmail", "emails", emailId, sendEmailRequest)).then(
                    function (sendResult) {
                        console.log("Email sent successfully!");

                        // --- Update Rev Status to "Issued" ---
                        Xrm.WebApi.updateRecord(entityLogicalName, entityId, {
                            "crbd5_revstatus": 769570000 // Replace with actual option set value for "Issued"
                        }).then(
                            function (updateRevStatusResult) {
                                console.log("Rev status updated to 'Issued' successfully.");

                                // Open the custom HTML web resource in a dialog
                                var pageInput = {
                                    pageType: "webresource",
                                    webresourceName: "new_emailNotification.html", // The name of your HTML web resource
                                };

                                var navigationOptions = {
                                    target: 2,
                                    width: { value: 400, unit: "px" },
                                    height: { value: 250, unit: "px" },
                                    title: "Pap Solutions Document Transmittal"
                                };

                                Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
                                    function success() {
                                        // Additional logic if needed after opening the dialog
                                    },
                                    function error(error) {
                                        console.log(error.message);
                                        Xrm.Navigation.openAlertDialog({
                                            title: "Pap Solutions Document Transmittal",
                                            text: "Error opening dialog: " + error.message
                                        });
                                    }
                                );
                            },
                            function (updateRevStatusError) {
                                console.error("Error updating rev status:", updateRevStatusError.message);
                                Xrm.Navigation.openAlertDialog({
                                    title: "Pap Solutions Document Transmittal",
                                    text: "Error updating Revision Status: " + updateRevStatusError.message
                                });
                            }
                        );
                    },
                    function (sendError) {
                        console.error("Error sending email:", sendError.message);
                        Xrm.Navigation.openAlertDialog({
                            title: "Pap Solutions Document Transmittal",
                            text: "Error sending email: " + sendError.message
                        });
                    }
                );
            },
            function (createError) {
                console.error("Error creating email record:", createError.message);
                Xrm.Navigation.openAlertDialog({
                    title: "Pap Solutions Document Transmittal",
                    text: "Error creating email record: " + createError.message
                });
            }
        );

    } catch (e) {
        console.error("Error in sendEmailWithDelay function:", e.message);
        Xrm.Navigation.openAlertDialog({
            title: "Pap Solutions Document Transmittal",
            text: "Exception: " + e.message
        });
    }
}