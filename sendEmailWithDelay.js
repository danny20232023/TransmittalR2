function sendEmailWithDelay() {
  try {
    // Ensure Xrm.Page is available
    if (!Xrm || !Xrm.Page || !Xrm.Page.data || !Xrm.Page.data.entity) {
      throw new Error("Xrm.Page.data.entity is unavailable.");
    }

    // Get the form context from Xrm.Page
    var formContext = Xrm.Page;

    // Get the record ID and entity logical name
    var entityId = formContext.data.entity.getId();
    var entityLogicalName = formContext.data.entity.getEntityName();

    // Update the record to trigger the email
    Xrm.WebApi.updateRecord(entityLogicalName, entityId, {
      "new_sendemailtrigger": new Date().toISOString()
    }).then(
      function success(result) {
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
            // Handle error if needed
            Xrm.Navigation.openAlertDialog({
              title: "Pap Solutions Document Transmittal",
              text: "Error: " + error.message
            });
          }
        );
      },
      function (error) {
        console.log(error.message);
        // Handle error if needed
        Xrm.Navigation.openAlertDialog({
          title: "Pap Solutions Document Transmittal",
          text: "Error: " + error.message
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
