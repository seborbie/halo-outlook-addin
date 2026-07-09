/* global Office, XMLHttpRequest, setTimeout, clearTimeout, window */
(function () {
  "use strict";

  var BACKGROUND_SESSION_STORAGE_KEY = "halo-auth-background-session-v1";
  var SEND_AUTO_ATTACH_URL = getOrigin() + "/api/halo/email/send-auto-attach";
  var SEND_AUTO_ATTACH_TIMEOUT_MS = 4500;

  function onHaloMessageSend(event) {
    try {
      readCurrentComposeEmail(function (readError, email) {
        if (readError || !email) {
          completeAllow(event);
          return;
        }

        sendAutoAttach(email, function (sendError, result) {
          if (sendError || !result) {
            completeAllow(event);
            return;
          }

          if (result.ok || result.status === "no-match" || result.status === "no-session") {
            completeAllow(event);
            return;
          }

          if (result.ticketNumber || result.ticketId) {
            completeWithHaloWarning(event, result);
            return;
          }

          completeAllow(event);
        });
      });
    } catch (error) {
      void error;
      completeAllow(event);
    }
  }

  function readCurrentComposeEmail(callback) {
    var item = Office.context.mailbox.item;
    if (!item || !isMessageItem(item) || !item.body) {
      callback(null, null);
      return;
    }

    var inReplyToMessageIds = [];
    if (item.inReplyTo) {
      inReplyToMessageIds.push(item.inReplyTo);
    }

    var conversationId = item.conversationId || "";
    if (!inReplyToMessageIds.length && !conversationId) {
      callback(null, null);
      return;
    }

    readMessageBody(item, function (bodyError, body) {
      if (bodyError || !body || (!body.bodyHtml && !body.bodyText)) {
        callback(null, null);
        return;
      }

      readSubject(item, function (subject) {
        readRecipients(item.to, function (to) {
          readRecipients(item.cc, function (cc) {
            readSender(item.from, function (from) {
              readItemId(item, function (itemId) {
                var userProfile = Office.context.mailbox.userProfile || {};
                var mailboxEmail = userProfile.emailAddress || "";

                callback(null, {
                  backgroundSessionId: getBackgroundSessionId(),
                  bodyHtml: body.bodyHtml,
                  bodyText: body.bodyText,
                  cc: cc,
                  conversationId: conversationId,
                  dateTimeCreated: new Date().toISOString(),
                  from: from || getUserProfileAddress(),
                  inReplyToMessageIds: inReplyToMessageIds,
                  internetHeaders: "",
                  internetMessageId: item.internetMessageId || "",
                  itemId: item.itemId || itemId,
                  mailboxEmail: mailboxEmail,
                  normalizedSubject: normalizeSubject(subject),
                  referenceMessageIds: [],
                  subject: subject,
                  timeZone: getClientTimeZone(),
                  to: to,
                });
              });
            });
          });
        });
      });
    });
  }

  function isMessageItem(item) {
    var itemType = item.itemType;
    var messageType = "";

    if (!itemType) {
      return true;
    }

    if (Office.MailboxEnums && Office.MailboxEnums.ItemType) {
      messageType = Office.MailboxEnums.ItemType.Message;
    }

    if (itemType === messageType) {
      return true;
    }

    if (String(itemType).toLowerCase() === "message") {
      return true;
    }

    return false;
  }

  function readMessageBody(item, callback) {
    if (!item.body || !item.body.getAsync) {
      callback(new Error("Could not read the compose body."));
      return;
    }

    item.body.getAsync(Office.CoercionType.Html, function (htmlResult) {
      if (isSucceeded(htmlResult)) {
        callback(null, {
          bodyHtml: htmlResult.value || "",
          bodyText: "",
        });
        return;
      }

      item.body.getAsync(Office.CoercionType.Text, function (textResult) {
        if (isSucceeded(textResult)) {
          callback(null, {
            bodyHtml: "",
            bodyText: textResult.value || "",
          });
          return;
        }

        callback(new Error(getAsyncErrorMessage(textResult, "Could not read the compose body.")));
      });
    });
  }

  function readSubject(item, callback) {
    if (typeof item.subject === "string") {
      callback(item.subject);
      return;
    }

    if (!item.subject || !item.subject.getAsync) {
      callback(item.normalizedSubject || "");
      return;
    }

    item.subject.getAsync(function (result) {
      if (isSucceeded(result)) {
        callback(result.value || "");
        return;
      }

      callback("");
    });
  }

  function readRecipients(value, callback) {
    if (isArray(value)) {
      callback(normalizeEmailAddressList(value));
      return;
    }

    if (!value || !value.getAsync) {
      callback([]);
      return;
    }

    value.getAsync(function (result) {
      if (isSucceeded(result)) {
        callback(normalizeEmailAddressList(result.value));
        return;
      }

      callback([]);
    });
  }

  function readSender(value, callback) {
    if (!value) {
      callback(null);
      return;
    }

    if (value.emailAddress || value.displayName) {
      callback(normalizeEmailAddress(value));
      return;
    }

    if (!value.getAsync) {
      callback(null);
      return;
    }

    value.getAsync(function (result) {
      if (isSucceeded(result)) {
        callback(normalizeEmailAddress(result.value));
        return;
      }

      callback(null);
    });
  }

  function readItemId(item, callback) {
    if (!item.getItemIdAsync) {
      callback("");
      return;
    }

    item.getItemIdAsync(function (result) {
      if (isSucceeded(result)) {
        callback(result.value || "");
        return;
      }

      callback("");
    });
  }

  function sendAutoAttach(email, callback) {
    var xhr = new XMLHttpRequest();
    var completed = false;
    var timeout = setTimeout(function () {
      finish(new Error("Halo send auto-attach timed out."));
    }, SEND_AUTO_ATTACH_TIMEOUT_MS);

    function finish(error, result) {
      if (completed) {
        return;
      }

      completed = true;
      clearTimeout(timeout);
      callback(error, result);
    }

    xhr.open("POST", SEND_AUTO_ATTACH_URL, true);
    xhr.withCredentials = true;
    xhr.setRequestHeader("Accept", "application/json");
    xhr.setRequestHeader("Content-Type", "application/json");

    xhr.onreadystatechange = function () {
      var parsed;

      if (xhr.readyState !== 4) {
        return;
      }

      try {
        parsed = JSON.parse(xhr.responseText || "{}");
      } catch (error) {
        void error;
        var successStatus = xhr.status >= 200 && xhr.status < 300;
        var fallbackStatus = "failed";
        if (successStatus) {
          fallbackStatus = "no-match";
        }

        parsed = {
          ok: successStatus,
          status: fallbackStatus,
        };
      }

      finish(null, parsed);
    };

    xhr.onerror = function () {
      finish(new Error("Halo send auto-attach request failed."));
    };

    xhr.send(JSON.stringify(email));
  }

  function getBackgroundSessionId() {
    var value = "";

    try {
      value = Office.context.roamingSettings.get(BACKGROUND_SESSION_STORAGE_KEY);
    } catch (error) {
      void error;
      value = "";
    }

    if (typeof value === "string") {
      return value;
    }

    return "";
  }

  function getOrigin() {
    if (window.location.origin) {
      return window.location.origin;
    }

    return window.location.protocol + "//" + window.location.host;
  }

  function getClientTimeZone() {
    try {
      return Intl.DateTimeFormat().resolvedOptions().timeZone || "";
    } catch (error) {
      void error;
      return "";
    }
  }

  function getUserProfileAddress() {
    var userProfile = Office.context.mailbox.userProfile;
    if (!userProfile || !userProfile.emailAddress) {
      return null;
    }

    return {
      displayName: userProfile.displayName || "",
      emailAddress: userProfile.emailAddress,
    };
  }

  function normalizeEmailAddressList(value) {
    var result = [];
    var index;
    var entry;

    if (!value) {
      return result;
    }

    for (index = 0; index < value.length; index += 1) {
      entry = normalizeEmailAddress(value[index]);
      if (entry) {
        result.push(entry);
      }
    }

    return result;
  }

  function normalizeEmailAddress(value) {
    if (!value) {
      return null;
    }

    return {
      displayName: value.displayName || "",
      emailAddress: value.emailAddress || "",
    };
  }

  function normalizeSubject(value) {
    return String(value || "")
      .replace(/^(\s*(re|fw|fwd)\s*:\s*)+/i, "")
      .trim();
  }

  function isArray(value) {
    if (Array.isArray) {
      return Array.isArray(value);
    }

    return Object.prototype.toString.call(value) === "[object Array]";
  }

  function isSucceeded(result) {
    var succeededStatus = "succeeded";

    if (Office.AsyncResultStatus && Office.AsyncResultStatus.Succeeded) {
      succeededStatus = Office.AsyncResultStatus.Succeeded;
    }

    return result && result.status === succeededStatus;
  }

  function getAsyncErrorMessage(result, fallback) {
    if (result && result.error && result.error.message) {
      return result.error.message;
    }

    return fallback;
  }

  function completeAllow(event) {
    event.completed({ allowEvent: true });
  }

  function completeWithHaloWarning(event, result) {
    var ticketLabel = result.ticketNumber || result.ticketId || "the mapped Halo ticket";
    event.completed({
      allowEvent: false,
      errorMessage:
        "Could not add this reply to Halo ticket " + ticketLabel + ". Send anyway or try again.",
    });
  }

  Office.actions.associate("onHaloMessageSend", onHaloMessageSend);
})();
