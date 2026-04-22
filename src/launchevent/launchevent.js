/* global Office OfficeRuntime console fetch */

function onMessageSendHandler(event) {
  console.log("--------------------------------TRIGGER-----------------------------");
  event.completed({
    allowEvent: true,
  });
}

function onAppointmentSendHandler(event) {
  console.log("Appointment send event triggered");
  event.completed({ allowEvent: true });
}

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
