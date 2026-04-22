/* global Office OfficeRuntime console fetch */

function onMessageSendHandler(event) {
  console.log("--------------------------------TRIGGER-----------------------------");
  Office.context.mailbox.item.loadCustomPropertiesAsync(
    { asyncContext: event },
    function (customPropsResult) {
      if (customPropsResult.status !== Office.AsyncResultStatus.Succeeded) {
        event.completed({
          allowEvent: false,
          errorMessage:
            "Die Partnerverknüpfung kann nicht hergestellt werden. Überprüfen Sie Ihre Netzwerkverbindung.",
        });
      }
      const customProps = customPropsResult.value;
      // In Outlook classic, the custom properties seem to lose the document payload, so we can't create the partner link.
      if (customProps.rawData.documentPayload) {
        createPartnerLinkAsync(
          customProps.rawData.linkedPartner,
          customProps.rawData.documentPayload,
          event
        );
      } else {
        event.completed({
          allowEvent: false,
          errorMessage:
            "Auf dem genutzten Outlook-Client kann der Partner nicht beim Erstellen einer Mail verknüpft werden. Bitte verknüpfen Sie den Partner, nachdem die Mail versendet wurde oder nutzen sie Outlook Web im Browser.",
        });
      }
    }
  );
}

function onAppointmentSendHandler(event) {
  console.log("Appointment send event triggered");
  event.completed({ allowEvent: true });
}

function createPartnerLinkAsync(partner, documentPayload, event) {
  const parsedPartner = JSON.parse(partner);
  const parsedPayload = JSON.parse(documentPayload);
  OfficeRuntime.auth
    .getAccessToken({ allowSignInPrompt: false, allowConsentPrompt: false })
    .then((accessToken) => {
      // webpack resolves the URL correctly but the launch event doesn't seem to derive from the dist folder in local development
      // Will switch between hard-coded and env variable URL based on local or dev/prod environment
      // const response = await fetch(`{{RESSOURCE_API}}`, {
      return fetch("https://localhost:7050/api/PartnerLink", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: "Bearer " + accessToken,
        },
        body: JSON.stringify({
          documentType: parsedPayload.documentType,
          documentId: parsedPayload.documentId,
          partnerId: parsedPartner.partner.key,
          partnerName: parsedPartner.partner.name,
          partnerAddress: parsedPartner.partner.addressMain,
          outlookUser: Office.context.mailbox.userProfile.emailAddress,
          aloaUser: "cn=50@AAAA000/ou=EBA_User/o=vkb", //TODO: get LTPA Username from AD
          state: parsedPayload.documentState,
          hasAttachment: parsedPayload.hasAttachment,
          lastOutlookUpdate: Office.context.mailbox.item.dateTimeModified,
          subject: parsedPayload.documentSubject,
          appointmentStartDatetime: parsedPayload.appointmentStartDatetime,
          mailCreatedDateTime:
            parsedPayload.documentType === 0 ? Office.context.mailbox.item.dateTimeCreated : null,
        }),
      });
    })
    .then((response) => {
      if (!response.ok) {
        throw new Error(response.status + ": " + response.statusText);
      }
    })
    .then(() => {
      event.completed({ allowEvent: true });
    })
    .catch(() => {
      event.completed({
        allowEvent: false,
        errorMessage: "Die Partnerverknüpfung kann nicht hergestellt werden.",
      });
    });
}

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
