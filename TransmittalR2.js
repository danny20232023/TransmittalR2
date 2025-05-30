function navigateToEmail() {
    // Retrieve the ID and name of the current Transmittal record
    var transmittalId = Xrm.Page.data.entity.getId();
    var transmittalNumber = Xrm.Page.data.entity.getPrimaryAttributeValue();
    var transmittalTitleAttr = Xrm.Page.data.entity.attributes.get("new_title");
    var transmittalProjectNumberAttr = Xrm.Page.data.entity.attributes.get("crbd5_projectnumber");

    var transmittalTitleValue = transmittalTitleAttr ? transmittalTitleAttr.getValue() : ""; 
    var transmittalProjectNumberValue = transmittalProjectNumberAttr ? transmittalProjectNumberAttr.getValue() : ""; 

    console.log("Transmittal ID:", transmittalId);
    console.log("Transmittal Number:", transmittalNumber);
    console.log("Transmittal Title:", transmittalTitleValue);
    console.log("Transmittal Project Number:", transmittalProjectNumberValue);

    // Retrieve ALL emails from the Project Transmittal CC table
    Xrm.WebApi.retrieveMultipleRecords("new_projecttransmittalcc", "?$select=new_email").then(
        function success(result) {
            var emailList = [];

            for (var i = 0; i < result.entities.length; i++) {
                var email = result.entities[i]["new_email"];
                if (email) {
                    emailList.push(email);
                }
            }

            var transmittalCC = emailList.join(", ");
            console.log("Transmittal CC:", transmittalCC);

            // Always create a new record
            console.log("Creating a new Sent Email Record.");

            var entityFormOptions = {
                entityName: "new_projecttransmittalsentemail"
            };

            var formParameters = {
                new_emailtransmittaluid: transmittalId,
                new_mail: transmittalTitleValue,
                new_cc: transmittalCC,
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
            console.error("Error retrieving Project Transmittal CC emails:", error.message);
        }
    );
}

// Important: Expose the function to the global scope so Power Apps can access it
window.navigateToEmail = navigateToEmail;
