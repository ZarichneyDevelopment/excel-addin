/* global Office */

// This add-in is for Excel. Keep the command handler Excel-safe (no Mailbox APIs).
Office.onReady(() => {
  // Office.js is ready.
});

function action(event: Office.AddinCommands.Event) {
  // Currently no ribbon commands are used; this exists to satisfy the manifest FunctionFile contract.
  console.log('[Budget Helper] Command invoked: action');
  event.completed();
}

Office.actions.associate('action', action);
