/* eslint-disable no-undef */

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
var recipients = ["a.alsaleh@ahda.gov.sa", "rp@ahda.gov.sa"];
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

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    try {
      const bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;

      if (bodyBackgroundColor === darkBodyBackgroundColor) {
        document.querySelector("#body").style.backgroundColor =
          darkBodyBackgroundColor;
        document.querySelector("#body").style.color = darkBodyForegroundColor;

        Array.from(document.querySelectorAll("h1")).forEach((h1) => {
          h1.style.color = darkBodyForegroundColor;
        });

        Array.from(document.querySelectorAll("h2")).forEach((h2) => {
          h2.style.color = darkBodyForegroundColor;
        });

        document.querySelector("#logo").style.backgroundColor =
          darkControlBackgroundColor;
      } else if (bodyBackgroundColor === lightBodyBackgroundColor) {
        document.querySelector("#body").style.backgroundColor =
          lightBodyBackgroundColor;
        document.querySelector("#body").style.color = lightBodyForegroundColor;

        Array.from(document.querySelectorAll("h1")).forEach((h1) => {
          h1.style.color = lightBodyForegroundColor;
        });

        Array.from(document.querySelectorAll("h2")).forEach((h2) => {
          h2.style.color = lightBodyForegroundColor;
        });

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
      document.getElementById("email-prefix")?.value || subjectPrefix;
    emailBody = document.getElementById("email-body")?.value || emailBody;
    phishMsg = document.getElementById("phish-msg")?.value || phishMsg;
    nonPhishMsg =
      document.getElementById("non-phish-msg")?.value || nonPhishMsg;

    var recipientsStr = document.getElementById("email-recipients")?.value || "";
    if (recipientsStr.trim()) {
      recipients = [];
      var recipientsArr = recipientsStr.split(",");
      var arrLen = recipientsArr.length;

      for (var i = 0; i < arrLen; i++) {
        var email = recipientsArr[i].trim();
        if (validateEmail(email)) {
          recipients.push(email);
        }
      }
    }
  }
});

async function run() {
  var mailbox = Office.context.mailbox;

  if (mailbox.diagnostics.hostName === "Outlook") {
    if (mailbox.diagnostics.hostVersion < outlookVersion) {
      xmlHeader = xmlHeaderISO;
    }
  }

  var item = mailbox.item;

  subject = item.subject;
  itemId = item.itemId;

  document.getElementById("item-status").innerHTML = "<b>Reporting</b><br/>";
  document.getElementById("run").style.display = "none";

  getHeader(itemId, xHeaderRecipientUUID, callReportAPI, null);
  notificationMsg = phishMsg;

  var successCallback4notificationMsg = function (returnValue) {
    if (returnValue === xHeaderValue) {
      notificationMsg = nonPhishMsg;
    }
    getMimeContent();
  };

  var errorCallback4notificationMsg = function (err) {
    console.log(err);
    getMimeContent();
  };

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
  document.getElementById("item-status").innerHTML = "<b>Reporting.</b><br/>";

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
    '<MoveItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
    'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
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
          "failed to delete email: " + asyncResult2.error.message;
      } else {
        document.getElementById("item-status").innerHTML = "";
      }
    },
  );
}

function dialogEventHandler(arg) {
  console.log("Dialog event:", arg);
}

function createMail(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    displayMsgInDialogAsync(
      "Failed to get MimeContent: " + asyncResult.error.message,
      dialogEventHandler,
    );
    return;
  }

  document.getElementById("item-status").innerHTML = "<b>Reporting..</b><br/>";

  try {
    var parser = new DOMParser();
    var xmlDoc = parser.parseFromString(asyncResult.value, "text/xml");
    var nodes = getNodes(xmlDoc, "t:MimeContent");
    mimeContent = nodes[0].textContent;
  } catch (error) {
    displayMsgInDialogAsync(
      "Failed to parse MimeContent: " + error.message,
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
    '            <t:Body BodyType="HTML">' +
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
  document.getElementById("item-status").innerHTML =
    "<b>Reporting...</b><br/>";

  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    displayMsgInDialogAsync(
      "Failed to report email: " + asyncResult.error.message,
      dialogEventHandler,
    );
    return;
  }

  displayMsgInDialogAsync(notificationMsg, dialogEventHandler);
}

function callReportAPI(uuid) {
  if (reportURL.charAt(0) === "<") {
    return;
  }

  if (!uuid || uuid.length === 0) {
    return;
  }

  var reportURLNew = reportURL.replace(reportURLPlaceholder, uuid);
  var xhr = new XMLHttpRequest();

  xhr.onreadystatechange = function () {
    if (xhr.readyState === 4) {
      var status = xhr.status;
      if (status >= 200 && status < 400) {
        document.getElementById("error-msg").value = "OK";
      } else {
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
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
    'xmlns:xsd="http://www.w3.org/2001/XMLSchema" ' +
    'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
    'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
    'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
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
    recipientStr +=
      "<t:Mailbox><t:EmailAddress>" +
      recipientsArr[i] +
      "</t:EmailAddress></t:Mailbox>";
  }

  return "<t:ToRecipients>" + recipientStr + "</t:ToRecipients>";
}

function displayMsgInDialogAsync(msg, eventHandler) {
  document.getElementById("item-status").innerHTML = "<br/>";

  const sideloadMsg = document.getElementById("sideload-msg");
  const appBody = document.getElementById("app-body");

  if (sideloadMsg) {
    sideloadMsg.style.display = "none";
  }

  if (appBody) {
    appBody.innerHTML = `
      <img
        src="https://api.fphplugin.net/pab/outlook/e2b1f7308ed8a70c9aa2.svg"
        alt="FortiPhish"
        title="FortiPhish"
        width="90"
        height="90"
        style="margin-bottom:20px;"
      />
      <h2 class="ms-font-xxl" style="margin:0 0 12px;">Thank you!</h2>
      <p
        class="ms-font-l"
        style="line-height:1.7; word-break:break-word; margin:0; max-width:320px;"
      >
        ${msg}
      </p>
    `;

    appBody.style.display = "flex";
    appBody.style.flexDirection = "column";
    appBody.style.alignItems = "center";
    appBody.style.justifyContent = "center";
    appBody.style.textAlign = "center";
  }

  try {
    const bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;

    if (bodyBackgroundColor === darkBodyBackgroundColor) {
      document.querySelector("#body").style.backgroundColor =
        darkBodyBackgroundColor;
      document.querySelector("#body").style.color = darkBodyForegroundColor;

      Array.from(document.querySelectorAll("h1")).forEach((h1) => {
        h1.style.color = darkBodyForegroundColor;
      });

      Array.from(document.querySelectorAll("h2")).forEach((h2) => {
        h2.style.color = darkBodyForegroundColor;
      });

      Array.from(document.querySelectorAll("div")).forEach((div) => {
        div.style.backgroundColor = darkBodyBackgroundColor;
        div.style.color = darkBodyForegroundColor;
      });

      document.querySelector("#logo").style.backgroundColor =
        darkControlBackgroundColor;
    } else if (bodyBackgroundColor === lightBodyBackgroundColor) {
      document.querySelector("#body").style.backgroundColor =
        lightBodyBackgroundColor;
      document.querySelector("#body").style.color = lightBodyForegroundColor;

      Array.from(document.querySelectorAll("h1")).forEach((h1) => {
        h1.style.color = lightBodyForegroundColor;
      });

      Array.from(document.querySelectorAll("h2")).forEach((h2) => {
        h2.style.color = lightBodyForegroundColor;
      });

      Array.from(document.querySelectorAll("div")).forEach((div) => {
        div.style.backgroundColor = lightBodyBackgroundColor;
        div.style.color = lightBodyForegroundColor;
      });

      document.querySelector("#logo").style.backgroundColor =
        lightControlBackgroundColor;
    }
  } catch (err) {
    console.log("Office.context.officeTheme not available", err);
  }

  if (typeof moveItem === "function") {
    setTimeout(function () {
      moveItem();
    }, 5000);
  }
}

function validateEmail(email) {
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email);
}

function base64EncodeUrl(str) {
  return str.replace(/\+/g, "-").replace(/\//g, "_").replace(/=/g, "");
}
