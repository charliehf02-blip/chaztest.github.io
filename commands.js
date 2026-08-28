/* Halo Request Forms - ribbon button action */
/* Opens the Halo self-service portal in a real popup window (Office Dialog API), */
/* not an iframe, so it isn't affected by Halo's X-Frame-Options: SAMEORIGIN header. */

Office.onReady(() => {
  // no-op: this file just needs to register the function below
});

const HALO_PORTAL_URL = "https://supportcentre.ahbl.ca/portal/";

function openHaloPortal(event) {
  Office.context.ui.displayDialogAsync(
    HALO_PORTAL_URL,
    { height: 80, width: 60, promptBeforeOpen: false },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(
          "Failed to open Halo portal dialog: " +
          asyncResult.error.code + " " + asyncResult.error.message
        );
      }
    }
  );

  // Tells Outlook the ribbon action has completed, so the button doesn't spin indefinitely.
  event.completed();
}

// Registers the function so the manifest's <FunctionName>openHaloPortal</FunctionName> can find it.
Office.actions.associate("openHaloPortal", openHaloPortal);
