/*global Office */

Office.onReady(() => {});
function action(event: Office.AddinCommands.Event) {
  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

Office.actions.associate("action", action);
