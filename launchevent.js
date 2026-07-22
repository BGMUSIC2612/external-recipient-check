/*
 * External Recipient Check — Smart Alerts OnMessageSend handler
 * Courmacs Legal Ltd
 *
 * Runs every time a user hits Send. If any To/Cc/Bcc recipient is outside
 * the domains listed in INTERNAL_DOMAINS, the send is paused and Outlook
 * shows a dialog listing the external addresses. "Send Anyway" = confirm.
 *
 * NOTE: this file must be plain JavaScript with no DOM access — in classic
 * Outlook for Windows it runs in a JS-only runtime (no window/document).
 */

// ---------------------------------------------------------------------------
// CONFIG — add every domain that counts as "internal"
// ---------------------------------------------------------------------------
var INTERNAL_DOMAINS = [
  "courmacslegal.co.uk",
  "401groupcouk.mail.onmicrosoft.com",
  "401group.co.uk",
  "courmacslegal.onmicrosoft.com",
  "401groupcouk.onmicrosoft.com",
  "companytriage.co.uk"
];

// If true, the email is allowed through when recipient lookup fails
// (network blip etc). If false, failures block the send. Fail-open is the
// sane default so the add-in can never stop the firm sending email.
var FAIL_OPEN = true;

// ---------------------------------------------------------------------------

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;

  Promise.all([
    getRecipients(item.to),
    getRecipients(item.cc),
    getRecipients(item.bcc)
  ])
    .then(function (results) {
      var all = [].concat(results[0], results[1], results[2]);
      var external = all.filter(isExternal);

      if (external.length === 0) {
        // All internal — send silently.
        event.completed({ allowEvent: true });
        return;
      }

      var lines = external.map(function (r) {
        var name = r.displayName && r.displayName !== r.emailAddress
          ? r.displayName + " — "
          : "";
        return "\u2022 " + name + r.emailAddress;
      });

      event.completed({
        allowEvent: false,
        errorMessage:
          "This email is going OUTSIDE Courmacs Legal:\n\n" +
          lines.join("\n") +
          "\n\nCheck each address is correct. Press 'Send Anyway' to confirm, " +
          "or go back and review the recipients."
      });
    })
    .catch(function () {
      event.completed({ allowEvent: FAIL_OPEN });
    });
}

function isExternal(recipient) {
  var address = (recipient.emailAddress || "").toLowerCase();
  var at = address.lastIndexOf("@");
  if (at === -1) {
    // Distribution lists etc. may not expose an SMTP address here.
    // Treat unresolvable entries as internal to avoid false alarms;
    // flip to `return true;` if you'd rather be strict.
    return false;
  }
  var domain = address.slice(at + 1);
  return INTERNAL_DOMAINS.indexOf(domain) === -1;
}

function getRecipients(field) {
  return new Promise(function (resolve, reject) {
    if (!field || !field.getAsync) {
      resolve([]);
      return;
    }
    field.getAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || []);
      } else {
        reject(result.error);
      }
    });
  });
}

// Register the handler name used in manifest.xml
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
