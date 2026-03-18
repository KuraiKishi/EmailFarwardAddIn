/* Custom Outlook add-in taskpane script focused on email forwarding/reporting */

/* global Office, document, DOMParser, window */

const CONFIG = {
  subjectPrefix: "[Email Forward] ",
  emailBody:
    "Please review the attached email. This message was forwarded from the Outlook add-in.",
  recipients: ["a.alsaleh@ahda.gov.sa"],
  successMessage: "تم تحويل الرسالة بنجاح.",
  failPrefix: "تعذر تحويل الرسالة: ",
  moveReportedItemToDeleted: false,
  moveItemFolderId: "deleteditems",
  minOutlookVersion: "15.0.4535.1004",
};

const XML_HEADER_UTF = '<?xml version="1.0" encoding="utf-8"?>';
const XML_HEADER_ISO = '<?xml version="1.0" encoding="iso-8859-1"?>';
let xmlHeader = XML_HEADER_UTF;

let itemId = "";
let mimeContent = "";
let originalSubject = "";

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  const runButton = document.getElementById("run");
  if (runButton) {
    runButton.onclick = run;
  }

  const sideloadMsg = document.getElementById("sideload-msg");
  const appBody = document.getElementById("app-body");
  if (sideloadMsg) sideloadMsg.style.display = "none";
  if (appBody) appBody.style.display = "flex";
});

export async function run() {
  const mailbox = Office.context.mailbox;
  if (!mailbox || !mailbox.item) {
    updateStatus("لا توجد رسالة مفتوحة.");
    return;
  }

  if (
    mailbox.diagnostics?.hostName === "Outlook" &&
    mailbox.diagnostics.hostVersion < CONFIG.minOutlookVersion
  ) {
    xmlHeader = XML_HEADER_ISO;
  }

  itemId = mailbox.item.itemId;
  originalSubject = mailbox.item.subject || "No subject";

  updateStatus("جاري تحويل الرسالة...");
  hideRunButton();
  getMimeContent();
}

function getMimeContent() {
  const request = addSoapHeader(
    "<m:GetItem>" +
      "<m:ItemShape>" +
      "<t:BaseShape>IdOnly</t:BaseShape>" +
      "<t:IncludeMimeContent>true</t:IncludeMimeContent>" +
      "</m:ItemShape>" +
      "<m:ItemIds>" +
      `<t:ItemId Id="${escapeXml(itemId)}" />` +
      "</m:ItemIds>" +
      "</m:GetItem>",
  );

  Office.context.mailbox.makeEwsRequestAsync(request, createForwardMessage);
}

function createForwardMessage(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    onFailure(asyncResult.error.message);
    return;
  }

  try {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(asyncResult.value, "text/xml");
    const mimeNodes = getNodes(xmlDoc, "t:MimeContent");
    if (!mimeNodes.length) {
      onFailure("MimeContent not found.");
      return;
    }
    mimeContent = mimeNodes[0].textContent;
  } catch (error) {
    onFailure(error.message);
    return;
  }

  const emlName = `${safeFileName(originalSubject)}.eml`;

  const xmlAttachments =
    "<t:Attachments>" +
    "<t:ItemAttachment>" +
    `<t:Name>${escapeXml(emlName)}</t:Name>` +
    "<t:IsInline>false</t:IsInline>" +
    "<t:Message>" +
    '<t:MimeContent CharacterSet="UTF-8">' +
    mimeContent +
    "</t:MimeContent>" +
    "</t:Message>" +
    "</t:ItemAttachment>" +
    "</t:Attachments>";

  const request = addSoapHeader(
    '<m:CreateItem MessageDisposition="SendOnly">' +
      '<m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems" /></m:SavedItemFolderId>' +
      "<m:Items>" +
      "<t:Message>" +
      `<t:Subject>${escapeXml(CONFIG.subjectPrefix + originalSubject)}</t:Subject>` +
      `<t:Body BodyType="HTML">${escapeXml(CONFIG.emailBody)}</t:Body>` +
      xmlAttachments +
      getToRecipients(CONFIG.recipients) +
      "</t:Message>" +
      "</m:Items>" +
      "</m:CreateItem>",
  );

  Office.context.mailbox.makeEwsRequestAsync(request, onForwardCompleted);
}

function onForwardCompleted(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    onFailure(asyncResult.error.message);
    return;
  }

  updateStatus(CONFIG.successMessage);
  showDialogMessage(CONFIG.successMessage);

  if (CONFIG.moveReportedItemToDeleted) {
    moveItem();
  }
}

function moveItem() {
  const request = addSoapHeader(
    '<MoveItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
      'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
      "<ToFolderId>" +
      `<t:DistinguishedFolderId Id="${escapeXml(CONFIG.moveItemFolderId)}"/>` +
      "</ToFolderId>" +
      "<ItemIds>" +
      `<t:ItemId Id="${escapeXml(itemId)}"/>` +
      "</ItemIds>" +
      "</MoveItem>",
  );

  Office.context.mailbox.makeEwsRequestAsync(request, () => {});
}

function showDialogMessage(message) {
  const encoded = base64EncodeUrl(window.btoa(message));
  const thanksUrl =
    "https://kuraikishi.github.io/EmailFarwardAddIn/thanks.html";

  Office.context.ui.displayDialogAsync(
    thanksUrl,
    { height: 40, width: 30, displayInIframe: false },
    () => {},
  );
}

function onFailure(message) {
  const fullMessage = CONFIG.failPrefix + message;
  updateStatus(fullMessage);
  showRunButton();
}

function updateStatus(message) {
  const statusNode = document.getElementById("item-status");
  if (statusNode) {
    statusNode.textContent = message;
  }
}

function hideRunButton() {
  const runButton = document.getElementById("run");
  if (runButton) runButton.style.display = "none";
}

function showRunButton() {
  const runButton = document.getElementById("run");
  if (runButton) runButton.style.display = "inline-block";
}

function getNodes(node, name) {
  const withoutNs = name.includes(":") ? name.split(":")[1] : name;
  let result = node.getElementsByTagName(name);
  if (!result || result.length === 0) {
    result = node.getElementsByTagName(withoutNs);
  }
  return result;
}

function addSoapHeader(request) {
  return (
    xmlHeader +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
    'xmlns:xsd="http://www.w3.org/2001/XMLSchema" ' +
    'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
    'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
    'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    "<soap:Header>" +
    '<RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    "</soap:Header>" +
    `<soap:Body>${request}</soap:Body>` +
    "</soap:Envelope>"
  );
}

function getToRecipients(recipients) {
  if (!recipients || !recipients.length) return "";
  const xml = recipients
    .filter(validateEmail)
    .map(
      (email) =>
        `<t:Mailbox><t:EmailAddress>${escapeXml(email)}</t:EmailAddress></t:Mailbox>`,
    )
    .join("");
  return `<t:ToRecipients>${xml}</t:ToRecipients>`;
}

function validateEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test((email || "").trim());
}

function base64EncodeUrl(str) {
  return str.replace(/\+/g, "-").replace(/\//g, "_").replace(/=/g, "");
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function safeFileName(value) {
  return String(value || "message")
    .replace(/[\\/:*?"<>|]/g, "_")
    .trim()
    .slice(0, 120);
}
