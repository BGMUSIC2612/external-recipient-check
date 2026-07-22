/* External Recipient Check — DEBUG BUILD v4 (adds Office.onReady)
 * Any failure now shows up in the send dialog instead of hanging.
 */

Office.onReady();

var INTERNAL_DOMAINS = [
  "courmacslegal.co.uk",
  "401groupcouk.mail.onmicrosoft.com",
  "401group.co.uk",
  "courmacslegal.onmicrosoft.com",
  "401groupcouk.onmicrosoft.com",
  "companytriage.co.uk"
];

function onMessageSendHandler(event) {
  var finished = false;

  function done(opts) {
    if (finished) return;
    finished = true;
    event.completed(opts);
  }

  // Watchdog: if nothing has completed within 15s, say so instead of hanging
  setTimeout(function () {
    done({
      allowEvent: false,
      errorMessage: "DEBUG v3: timed out reading recipients (getAsync never returned)."
    });
  }, 15000);

  try {
    var item = Office.context.mailbox.item;
    if (!item) {
      done({ allowEvent: false, errorMessage: "DEBUG v3: mailbox item unavailable." });
      return;
    }

    Promise.all([
      getRecipients(item.to),
      getRecipients(item.cc),
      getRecipients(item.bcc)
    ])
      .then(function (results) {
        var all = [].concat(results[0], results[1], results[2]);
        var external = all.filter(isExternal);

        if (external.length === 0) {
          done({ allowEvent: true });
          return;
        }

        var lines = external.map(function (r) {
          var name = r.displayName && r.displayName !== r.emailAddress
            ? r.displayName + " — " : "";
          return "\u2022 " + name + r.emailAddress;
        });

        done({
          allowEvent: false,
          errorMessage:
            "This email is going OUTSIDE Courmacs Legal:\n\n" +
            lines.join("\n") +
            "\n\nCheck each address is correct. Press 'Send Anyway' to confirm, " +
            "or go back and review the recipients."
        });
      })
      .catch(function (err) {
        done({
          allowEvent: false,
          errorMessage: "DEBUG v3: recipient lookup failed — " +
            (err && err.message ? err.message : JSON.stringify(err))
        });
      });
  } catch (e) {
    done({
      allowEvent: false,
      errorMessage: "DEBUG v3: handler exception — " + (e && e.message ? e.message : String(e))
    });
  }
}

function isExternal(recipient) {
  var address = (recipient.emailAddress || "").toLowerCase();
  var at = address.lastIndexOf("@");
  if (at === -1) return false;
  return INTERNAL_DOMAINS.indexOf(address.slice(at + 1)) === -1;
}

function getRecipients(field) {
  return new Promise(function (resolve, reject) {
    if (!field || !field.getAsync) { resolve([]); return; }
    field.getAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || []);
      } else {
        reject(result.error || new Error("getAsync failed"));
      }
    });
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
