/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var xmlHeaderUTF = '<?xml version="1.0" encoding="utf-8"?>';
var xmlHeaderISO = '<?xml version="1.0" encoding="iso-8859-1"?>';
var outlookVersion = "15.0.4535.1004";
var xmlHeader = xmlHeaderUTF;

var reportURL = "<RECIPIENT_UUID>";
var reportURLPlaceholder = "<RECIPIENT_UUID>";
var xHeaderName = "X-PHISHTEST";
var xHeaderValue = "FortiPhish";
var xHeaderRecipientUUID = "X-FORTIPHISH-RECIPIENT-UUID";

var itemId = "";
var mimeContent = "";
var subject = "Reported Phishing Email";
var recipients = ["a.alsaleh@ahda.gov.sa, rp@ahda.gov.sa"];
var subjectPrefix = "[Phish Alert]";
var emailBody =
  "I would like to report the email enclosed as a phishing email.";
var moveItemFolderId = "deleteditems";
var phishMsg =
  "Thank you for reporting this email. Because of people like you, our company is more secure!";
var nonPhishMsg =
  "Congratulations, you successfully identified a simulated phishing email!";
var notificationMsg = "";

const darkBodyBackgroundColor = "#212121";
const darkBodyForegroundColor = "#FFFFFF";
const darkControlBackgroundColor = "#292929";
const darkControlForegroundColor = "#D6D6D6";

const lightBodyBackgroundColor = "#FAF9F8";
const lightBodyForegroundColor = "#242424";
const lightControlBackgroundColor = "#FFFFFF";
const lightControlForegroundColor = "#424242";

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    try {
      const bodyBackgroundColor =
        Office.context.officeTheme.bodyBackgroundColor;
      console.log(
        "Office.context.officeTheme.bodyBackgroundColor",
        bodyBackgroundColor,
      ); // for debug
      if (bodyBackgroundColor === darkBodyBackgroundColor) {
        // Outlook dark theme

        document.querySelector("#body").style.backgroundColor =
          darkBodyBackgroundColor;
        document.querySelector("#body").style.color = darkBodyForegroundColor;
        Array.from(document.querySelectorAll("h1")).map(
          (h1) => (h1.style.color = darkBodyForegroundColor),
        );

        Array.from(document.querySelectorAll("h2")).map(
          (h2) => (h2.style.color = darkBodyForegroundColor),
        );
        document.querySelector("#logo").style.backgroundColor =
          darkControlBackgroundColor;
      } else if (bodyBackgroundColor === lightBodyBackgroundColor) {
        // Outlook light theme
        document.querySelector("#body").style.backgroundColor =
          lightBodyBackgroundColor;
        document.querySelector("#body").style.color = lightBodyForegroundColor;
        Array.from(document.querySelectorAll("h1")).map(
          (h1) => (h1.style.color = lightBodyForegroundColor),
        );
        Array.from(document.querySelectorAll("h2")).map(
          (h2) => (h2.style.color = lightBodyForegroundColor),
        );
        document.querySelector("#logo").style.backgroundColor =
          lightControlBackgroundColor;
      }
    } catch (err) {
      console.log("Office.context.officeTheme not available", err);
    }
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    subjectPrefix =
      document.getElementById("email-prefix").value || subjectPrefix;
    emailBody = document.getElementById("email-body").value || emailBody;
    phishMsg = document.getElementById("phish-msg").value || phishMsg;
    nonPhishMsg = document.getElementById("non-phish-msg").value || nonPhishMsg;

    var recipientsStr = document.getElementById("email-recipients").value;
    var recipientsArr = recipientsStr.split(",");
    var arrLen = recipientsArr.length;
    for (var i = 0; i < arrLen; i++) {
      if (validateEmail(recipientsArr[i])) {
        recipients.push(recipientsArr[i]);
      }
    }
  }
});

// user clicked report button
async function run() {
  var mailbox = Office.context.mailbox;

  if (mailbox.diagnostics.hostName === "Outlook") {
    if (mailbox.diagnostics.hostVersion < outlookVersion) {
      xmlHeader = xmlHeaderISO; // backward compatibility headers
    }
  }

  // Get reference to current message
  var item = mailbox.item;

  // Set subject
  subject = item.subject;

  // Get current selected mail message Id
  itemId = item.itemId;

  // Update status
  document.getElementById("item-status").innerHTML = "<b>Reporting</b> <br/>";
  document.getElementById("run").style.display = "none";

  // call report API
  getHeader(itemId, xHeaderRecipientUUID, callReportAPI, null);
  // set default value [dialog message]
  notificationMsg = phishMsg;

  var successCallback4notificationMsg = function (returnValue) {
    console.log(returnValue, xHeaderValue, returnValue === xHeaderValue);
    if (returnValue === xHeaderValue) {
      notificationMsg = nonPhishMsg;
    }
    getMimeContent();
  };
  var errorCallback4notificationMsg = function (err) {
    console.log(err);
    getMimeContent();
  };

  // Get x-header from current selected mail message
  getHeader(
    itemId,
    xHeaderName,
    successCallback4notificationMsg,
    errorCallback4notificationMsg,
  );
}

function getHeader(itemId, headerName, successCallback, errorCallback) {
  var headerType = "String";
  var request =
    "   <m:GetItem>" +
    "       <m:ItemShape>" +
    "           <t:BaseShape>IdOnly</t:BaseShape>" +
    "           <t:AdditionalProperties>" +
    '               <t:ExtendedFieldURI DistinguishedPropertySetId="InternetHeaders" PropertyName="' +
    headerName +
    '" PropertyType="' +
    headerType +
    '" />' +
    "           </t:AdditionalProperties>" +
    "       </m:ItemShape>" +
    "       <m:ItemIds>" +
    '           <t:ItemId Id="' +
    itemId +
    '" />' +
    "       </m:ItemIds>" +
    "   </m:GetItem>";

  request = addSoapHeader(request);

  Office.context.mailbox.makeEwsRequestAsync(request, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      if (errorCallback != null) {
        errorCallback(asyncResult.error.message);
      }
      return;
    }

    var returnValue = "";
    try {
      // Get Header
      var parser = new DOMParser();
      var xmlDoc = parser.parseFromString(asyncResult.value, "text/xml");

      if (xmlDoc != null) {
        var nodes = getNodes(xmlDoc, "t:ExtendedProperty");
        for (var idx = 0; idx < nodes.length; idx++) {
          var value = nodes[idx];
          var nodeName = getNodes(value, "t:ExtendedFieldURI")[0].getAttribute(
            "PropertyName",
          );
          var nodeValue = getNodes(value, "t:Value")[0].textContent;
          if (nodeName === headerName) {
            returnValue = nodeValue;
            break;
          }
        }
        if (successCallback != null) {
          successCallback(returnValue);
        }
      }
    } catch (error) {
      if (errorCallback != null) {
        errorCallback(error.message);
      }
    }
  });
}

function getMimeContent() {
  // step 1
  document.getElementById("item-status").innerHTML = "<b>Reporting.</b> <br/>";

  // Get Mime content from current selected mail message
  var request_MimeContent =
    "       <m:GetItem>" +
    "           <m:ItemShape>" +
    "               <t:BaseShape>IdOnly</t:BaseShape>" +
    "               <t:IncludeMimeContent>true</t:IncludeMimeContent>" +
    "           </m:ItemShape>" +
    "           <m:ItemIds>" +
    '               <t:ItemId Id="' +
    itemId +
    '" />' +
    "           </m:ItemIds>" +
    "       </m:GetItem>";
  request_MimeContent = addSoapHeader(request_MimeContent);

  Office.context.mailbox.makeEwsRequestAsync(request_MimeContent, createMail);
}

function moveItem() {
  var request_MoveItem =
    '<MoveItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
    '          xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    "    <ToFolderId>" +
    '        <t:DistinguishedFolderId Id="' +
    moveItemFolderId +
    '"/>' +
    "    </ToFolderId>" +
    "    <ItemIds>" +
    '        <t:ItemId Id="' +
    itemId +
    '"/>' +
    "    </ItemIds>" +
    "</MoveItem>";
  request_MoveItem = addSoapHeader(request_MoveItem);

  Office.context.mailbox.makeEwsRequestAsync(
    request_MoveItem,
    function (asyncResult2) {
      if (asyncResult2.status === Office.AsyncResultStatus.Failed) {
        console.log("request_MoveItem ", asyncResult2.error.message);
        document.getElementById("item-status").innerHTML =
          "failed to delete email:" + asyncResult2.error.message;
      } else {
        document.getElementById("item-status").innerHTML = "";
      }
    },
  );
}

function dialogEventHandler(arg) {
  // https://learn.microsoft.com/en-us/office/dev/add-ins/reference/javascript-api-for-office-error-codes

  switch (arg.error) {
    case 12002:
      console.log("Cannot load URL, no such page or bad URL syntax");
      moveItem();
      break;
    case 12003:
      console.log("HTTPS is required");
      moveItem();
      break;
    case 12004:
      console.log(
        "The domain of the URL passed to displayDialogAsync is not trusted.",
      );
      moveItem();
      break;
    case 12005:
      console.log(
        "The URL passed to displayDialogAsync uses the HTTP protocol. HTTPS is required. ",
      );
      moveItem();
      break;
    case 12006:
      // The dialog was closed, typically because user pressed X button
      console.log("Dialog closed by user");
      moveItem();
      break;
    case 12007:
      console.log("A dialog box is already opened from this host window.");
      moveItem();
      break;
    case 12009:
      console.log("The user chose to ignore the dialog box.");
      moveItem();
      break;
    case 12011:
      console.log(
        "The user's browser is configured in a way that blocks popups.",
      );
      moveItem();
      break;
    default:
      console.log("Undefined error in dialog window", arg);
      break;
  }
}

function createMail(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    displayMsgInDialogAsync(
      "failed to get MimeContent:" + asyncResult.error.message,
      dialogEventHandler,
    );
    return;
  }
  // step 2
  document.getElementById("item-status").innerHTML = "<b>Reporting..</b> <br/>";

  try {
    // Get MimeContent
    var parser = new DOMParser();
    var xmlDoc = parser.parseFromString(asyncResult.value, "text/xml");
    var nodes = getNodes(xmlDoc, "t:MimeContent");
    mimeContent = nodes[0].textContent;
  } catch (error) {
    displayMsgInDialogAsync(
      "failed to parse MimeContent:" + error.message,
      dialogEventHandler,
    );
    return;
  }

  var xmlAttachments =
    "<t:Attachments>" +
    "<t:ItemAttachment>" +
    "<t:Name>" +
    subject +
    ".eml</t:Name>" +
    "<t:IsInline>false</t:IsInline>" +
    "<t:Message>" +
    '<t:MimeContent CharacterSet="UTF-8">' +
    mimeContent +
    "</t:MimeContent>" +
    "</t:Message>" +
    "</t:ItemAttachment>" +
    "</t:Attachments>";

  var xmlSavedFolder =
    '<m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems" /></m:SavedItemFolderId>';

  var request_Email =
    '<m:CreateItem MessageDisposition="SendOnly">' +
    xmlSavedFolder +
    "    <m:Items>" +
    "        <t:Message>" +
    "            <t:Subject>" +
    subjectPrefix +
    subject +
    "</t:Subject>" +
    '          <t:Body BodyType="HTML">' +
    emailBody +
    "</t:Body>" +
    xmlAttachments +
    getToRecipients(recipients) +
    "        </t:Message>" +
    "    </m:Items>" +
    "</m:CreateItem>";
  request_Email = addSoapHeader(request_Email);
  Office.context.mailbox.makeEwsRequestAsync(request_Email, reqDisplayDialog);
}

function reqDisplayDialog(asyncResult) {
  // step 3
  document.getElementById("item-status").innerHTML =
    "<b>Reporting...</b> <br/>";

  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    displayMsgInDialogAsync(
      "failed to report email:" + asyncResult.error.message,
      dialogEventHandler,
    );
    return;
  }

  // show popup notification here
  displayMsgInDialogAsync(notificationMsg, dialogEventHandler);
}

function callReportAPI(uuid) {
  if (reportURL.charAt(0) === "<") {
    return;
  }

  if (uuid.length === 0) {
    return;
  }

  var reportURLNew = reportURL.replace(reportURLPlaceholder, uuid);
  var xhr = new XMLHttpRequest();

  xhr.onreadystatechange = function () {
    if (xhr.readyState === 4) {
      var status = xhr.status;
      if (status >= 200 && status < 400) {
        // console.log(xhr.responseText);
        document.getElementById("error-msg").value = "OK";
      } else {
        // error
        document.getElementById("error-msg").value =
          xhr.statusText + "<-_->" + xhr.responseText;
      }
    }
  };

  xhr.open("GET", reportURLNew, true);
  xhr.send();
}

function getNodes(node, elementNameWithNS) {
  var elementWithoutNS = elementNameWithNS.substring(
    elementNameWithNS.indexOf(":") + 1,
  );
  var retVal = node.getElementsByTagName(elementNameWithNS);
  if (retVal == null || retVal.length === 0) {
    retVal = node.getElementsByTagName(elementWithoutNS);
  }
  return retVal;
}

function addSoapHeader(request) {
  var result =
    xmlHeader +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    "   <soap:Header>" +
    '       <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    "   </soap:Header>" +
    "   <soap:Body>" +
    request +
    "</soap:Body>" +
    "</soap:Envelope>";
  return result;
}

function getToRecipients(recipientsArr) {
  var arrLen = recipientsArr.length;
  if (arrLen === 0) {
    return "";
  }

  var recipientStr = "";
  for (var i = 0; i < arrLen; i++) {
    recipientStr =
      recipientStr +
      "<t:Mailbox><t:EmailAddress>" +
      recipientsArr[i] +
      "</t:EmailAddress></t:Mailbox>";
  }

  return "<t:ToRecipients>" + recipientStr + "</t:ToRecipients>";
}

function displayMsgInDialogAsync(msg, eventHandler) {
  var encoded = window.btoa(msg);
  encoded = base64EncodeUrl(encoded);

  // clear text
  document.getElementById("item-status").innerHTML = "<br/>";

  Office.context.ui.displayDialogAsync(
    window.location.origin + "/pab/outlook/thanks.html#" + encoded,
    { height: 60, width: 40, displayInIframe: false },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        eventHandler({ error: asyncResult.error.code });
        return;
      }
      const dialog = asyncResult.value;
      dialog.addEventHandler(
        Office.EventType.DialogEventReceived,
        eventHandler,
      );
    },
  );
  try {
    const bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
    console.log(
      "Office.context.officeTheme.bodyBackgroundColor",
      bodyBackgroundColor,
    ); // for debug
    if (bodyBackgroundColor === darkBodyBackgroundColor) {
      // Outlook dark theme

      document.querySelector("#body").style.backgroundColor =
        darkBodyBackgroundColor;
      document.querySelector("#body").style.color = darkBodyForegroundColor;
      Array.from(document.querySelectorAll("h1")).map(
        (h1) => (h1.style.color = darkBodyForegroundColor),
      );
      Array.from(document.querySelectorAll("div")).map((div) => {
        div.style.backgroundColor = darkBodyBackgroundColor;
        div.style.color = darkBodyForegroundColor;
      });
      Array.from(document.querySelectorAll("h2")).map(
        (h2) => (h2.style.color = darkBodyForegroundColor),
      );
      document.querySelector("#logo").style.backgroundColor =
        darkControlBackgroundColor;
    } else if (bodyBackgroundColor === lightBodyBackgroundColor) {
      // Outlook light theme
      document.querySelector("#body").style.backgroundColor =
        lightBodyBackgroundColor;
      document.querySelector("#body").style.color = lightBodyForegroundColor;
      Array.from(document.querySelectorAll("h1")).map(
        (h1) => (h1.style.color = lightBodyForegroundColor),
      );
      Array.from(document.querySelectorAll("h2")).map(
        (h2) => (h2.style.color = lightBodyForegroundColor),
      );
      Array.from(document.querySelectorAll("div")).map((div) => {
        div.style.backgroundColor = lightBodyBackgroundColor;
        div.style.color = lightBodyForegroundColor;
      });
      document.querySelector("#logo").style.backgroundColor =
        lightControlBackgroundColor;
    }
  } catch (err) {
    console.log("Office.context.officeTheme not available", err);
  }
}

function validateEmail(email) {
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email);
}

// base64-encoded to base64-url-encoded
function base64EncodeUrl(str) {
  return str.replace(/\+/g, "-").replace(/\//g, "_").replace(/=/g, "");
}
