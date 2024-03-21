/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("dialog-button").onclick = testingNewThing;
    document.getElementById("open-dialog").onclick = openDialog;
  }
});

export async function testingNewThing(){
  window.location.replace("https://localhost:3000/form.html");
}

// export async function openDialogue() {
//   const item = Office.context.mailbox.item;
//   document.getElementById("item-subject").innerHTML = "<b>Mail Created at:</b> <br/>" + item.dateTimeCreated;
//   let url = new URI('../src/form/form.html').absoluteTo(window.location).toString();
//   const dialogOptions = { width: 200, height: 400, displayInIframe: false };
//   Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
//     settingsDialog = result.value;
//     const item = Office.context.mailbox.item;
//     document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + JSON.stringify(item.body);
//     settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, receiveMessage);
//     settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, dialogClosed);
//   });
// }
let dialog = null;
function openDialog() {
  Office.context.ui.displayDialogAsync(
    'https://localhost:3000/popup.html',
    { height: 45, width: 55 },

    function (result) {
      dialog = result.value;
      dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
  );
}

function processMessage(arg) {
  document.getElementById("user-name").innerHTML = arg.message;
  dialog.close();
}

export async function run() {
  // Get a reference to the current message
  const item = Office.context.mailbox.item;

  // Write message property value to the task pane
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
}