/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
// https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough?tabs=xmlmanifest
// https://learn.microsoft.com/en-us/samples/officedev/pnp-officeaddins/using-storage-to-share-data-between-ui-less-custom-functions-and-the-task-pane/

async function StoreValue(key, value) {
  console.log("OfficeRuntime.storage: ",OfficeRuntime.storage);
  return OfficeRuntime.storage.setItem(key, value).then(
    function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
    },
    function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
    }
  );
}

async function GetValue(key) {
  return OfficeRuntime.storage.getItem(key);
}

async function onMessageSendHandler(event) {
  await StoreValue("key1", "value1");
  Office.context.mailbox.item.body.getAsync("text", { asyncContext: event }, getBodyCallback);
}

function getBodyCallback(asyncResult) {
  let event = asyncResult.asyncContext;
  let body = "";
  if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
    body = asyncResult.value;
  } else {
    let message = "Failed to get body text";
    console.error(message);
    event.completed({ allowEvent: false, errorMessage: message });
    return;
  }

  let matches = hasMatches(body);
  if (matches) {
    Office.context.mailbox.item.getAttachmentsAsync({ asyncContext: event }, getAttachmentsCallback);
  } else {
    event.completed({ allowEvent: true });
  }
}

function hasMatches(body) {
  if (body == null || body == "") {
    return false;
  }

  const arrayOfTerms = ["send", "picture", "document", "attachment"];
  for (let index = 0; index < arrayOfTerms.length; index++) {
    const term = arrayOfTerms[index].trim();
    const regex = RegExp(term, "i");
    if (regex.test(body)) {
      return true;
    }
  }

  return false;
}

function getAttachmentsCallback(asyncResult) {
  // const storageValue = await GetValue("key1");
  // console.log("Value for key1 is ", storageValue);
  let event = asyncResult.asyncContext;
  if (asyncResult.value.length > 0) {
    for (let i = 0; i < asyncResult.value.length; i++) {
      if (asyncResult.value[i].isInline == false) {
        event.completed({ allowEvent: true });
        return;
      }
    }

    event.completed({ allowEvent: false, errorMessage: "Looks like you forgot to include an attachment?" });
  } else {
    event.completed({ allowEvent: false, errorMessage: "Looks like you're forgetting to include an attachment?" });
  }
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
// if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
// }
