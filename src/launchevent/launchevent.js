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
