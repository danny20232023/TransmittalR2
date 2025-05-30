function navigateToEmail() {
    // Retrieve the ID and name of the current Transmittal record
    var transmittalId = Xrm.Page.data.entity.getId();
    var transmittalNumber = Xrm.Page.data.entity.getPrimaryAttributeValue();
    var transmittalTitle = Xrm.Page.data.entity.attributes.get("new_title");
    var transmittalProjectNumber = Xrm.Page.data.entity.attributes.get("crbd5_projectnumber");
    var transmittalTitleValue = transmittalTitle ? transmittalTitle.getValue() : ""; 
    var transmittalProjectNumberValue = transmittalProjectNumber ? transmittalProjectNumber.getValue() : ""; 
    
    console.log("Transmittal ID:", transmittalId);
    console.log("Transmittal Number:", transmittalNumber);
    console.log("Transmittal Title:", transmittalTitleValue);
    console.log("Transmittal Project Number:", transmittalProjectNumberValue);
    console.log("Transmittal CC:", transmittalCC);

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

            var entityFormOptions = {};
            entityFormOptions["entityName"] = "new_projecttransmittalsentemail";

    // Always create a new record
    console.log("Creating a new Sent Email Record.");

    var entityFormOptions = {};
    entityFormOptions["entityName"] = "new_projecttransmittalsentemail";

    var formParameters = {};
    formParameters["new_emailtransmittaluid"] = transmittalId;
    formParameters["new_mail"] = transmittalTitleValue;
    formParameters["new_cc"] = transmittalCC;
    formParameters["new_subject"] = "Transmittal " + transmittalNumber + " for " + transmittalTitleValue; 
    formParameters["new_body"] = 
        "<div style='font-family: Aptos, sans-serif; font-size: 15px;'>" +
        "<p style='margin-bottom: 10px;'>Hi Client<strong style='color: blue;'></strong>,</p>" +
        "<p style='margin-bottom: 5px;'>Please find attached transmitted documents for <strong style='color: green;'>" + transmittalTitleValue + "</strong>.</p>" +
        "<p style='margin-bottom: 10px;'>Can you please acknowledge receipt and return to <a href='mailto:documentcontrol@papsolutions.com.au'>documentcontrol@papsolutions.com.au</a>.</p>" +
        "<p style='margin-bottom: 10px;'>Kind Regards,</p>" +
        "<p><strong>Pap Solutions Pty Ltd Document Control</strong></p>" +
        "</div>";

    console.log("Opening new Sent Email Form with parameters:", formParameters);
     Xrm.Navigation.openForm(entityFormOptions, formParameters);
}
