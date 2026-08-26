/* global Office, XMLHttpRequest, setTimeout, clearTimeout, window, atob, btoa */
(function () {
  "use strict";

  var BACKGROUND_SESSION_STORAGE_KEY = "halo-auth-background-session-v1";
  var COMPOSE_ATTACH_STORAGE_KEY = "halo.composeAttach.v1";
  var COMPOSE_EMAIL_ATTACHMENT_STORAGE_KEY = "halo.emailAttachmentPrefetch.v2";
  var MAPPING_RECOVERY_STORAGE_KEY = "halo.mappingRecovery.v1";
  var COMPOSE_ATTACH_HEADER_NAME = "X-Halo-Compose-Id";
  var COMPOSE_CUSTOM_PROPERTY_NAME = "cecp-55bbcff2-8191-4411-aec6-f9d2f9b4b5e8";
  var SEND_AUTO_ATTACH_URL = "__HALO_PUBLIC_BASE_URL__/api/halo/email/send-auto-attach";
  var RECOVER_MAPPING_URL = "__HALO_PUBLIC_BASE_URL__/api/halo/email/recover-mapping";
  var SEND_DIAGNOSTIC_URL = "__HALO_PUBLIC_BASE_URL__/api/diagnostics/send-event";
  var SEND_EXPLICIT_ATTACH_BASE_URL = "__HALO_PUBLIC_BASE_URL__/api/halo/tickets/";
  var SEND_CREATE_TICKET_BASE_URL = "__HALO_PUBLIC_BASE_URL__/api/halo/ticket-creation/intents/";
  var EMAIL_ATTACHMENT_PREFETCH_START_URL =
    "__HALO_PUBLIC_BASE_URL__/api/halo/email-attachments/prefetch/start";
  var EMAIL_ATTACHMENT_PREFETCH_BASE_URL =
    "__HALO_PUBLIC_BASE_URL__/api/halo/email-attachments/prefetch/";
  var SEND_REQUEST_TIMEOUT_MS = 20000;
  var SEND_EVENT_TIMEOUT_MS = 25000;
  var MAX_INLINE_IMAGE_BYTES = 2 * 1024 * 1024;
  var MAX_INLINE_IMAGE_TOTAL_BYTES = 5 * 1024 * 1024;
  var MAX_EMAIL_ATTACHMENTS = 20;
  var MAX_EMAIL_ATTACHMENT_BYTES = 25 * 1024 * 1024;
  var MAX_EMAIL_ATTACHMENT_TOTAL_BYTES = 50 * 1024 * 1024;
  var EMAIL_ATTACHMENT_READ_RETRY_DELAYS_MS = [150, 350];
  var EMAIL_ATTACHMENT_CONTENT_READ_TIMEOUT_MS = 15000;
  var EMAIL_ATTACHMENT_EVENT_BUDGET_MS = 4 * 60 * 1000;
  var RECOVERY_TIMEOUT_MS = 6000;
  var recoveryFlights = {};

  function onHaloMessageSend(event) {
    var sendStartedAt = Date.now();
    var selectedMarker = null;
    var automaticAttachmentWarning = null;
    var watchdog = null;
    var complete = createSendEventCompletion(event, function () {
      clearTimeout(watchdog);
    });
    watchdog = setTimeout(function () {
      sendRuntimeDiagnostic("watchdog-fired", sendStartedAt, { outcome: "timeout" });
      if (selectedMarker) {
        completeWithHaloWarning(complete, selectedMarker, "timeout");
        return;
      }
      if (automaticAttachmentWarning) {
        completeWithHaloWarning(complete, automaticAttachmentWarning, "timeout");
        return;
      }

      completeAllow(complete);
    }, SEND_EVENT_TIMEOUT_MS);

    try {
      sendRuntimeDiagnostic("event-start", sendStartedAt, { outcome: "started" });
      readComposeAttachMarker(function (markerError, marker) {
        if (markerError) {
          sendRuntimeDiagnostic("marker-read", sendStartedAt, { outcome: "failed" });
          completeAllow(complete);
          return;
        }

        selectedMarker = marker;
        if (selectedMarker) {
          var markerOutcome = "explicit-ticket";
          if (isCreateTicketMarker(selectedMarker)) {
            markerOutcome = "create-ticket";
          }
          sendRuntimeDiagnostic("marker-read", sendStartedAt, {
            hasBackgroundSession: Boolean(selectedMarker.backgroundSessionId),
            hasMailboxIdentity: Boolean(selectedMarker.mailboxEmail),
            outcome: markerOutcome,
          });
          if (isCreateTicketMarker(selectedMarker)) {
            sendComposeTicketCreation(selectedMarker, complete, sendStartedAt);
            return;
          }
          sendExplicitComposeAttachment(selectedMarker, complete, sendStartedAt);
          return;
        }

        sendRuntimeDiagnostic("marker-read", sendStartedAt, { outcome: "automatic" });
        sendRuntimeDiagnostic("compose-read-start", sendStartedAt, { outcome: "started" });
        readCurrentComposeEmail(false, function (readError, email) {
          if (readError || !email) {
            sendRuntimeDiagnostic("compose-read-complete", sendStartedAt, {
              outcome: readError ? "failed" : "allowed-no-email",
            });
            sendRuntimeDiagnostic("event-complete", sendStartedAt, { outcome: "allowed-no-email" });
            completeAllow(complete);
            return;
          }

          sendRuntimeDiagnostic("compose-read-complete", sendStartedAt, {
            hasBackgroundSession: Boolean(email.backgroundSessionId),
            hasMailboxIdentity: Boolean(email.mailboxEmail),
            outcome: "ok",
          });
          recoverAutomaticMapping(Office.context.mailbox.item, email, function (recovery) {
            if (!recovery || recovery.status !== "matched") {
              sendRuntimeDiagnostic("event-complete", sendStartedAt, {
                outcome: "allowed-auto-result",
              });
              completeAllow(complete);
              return;
            }
            applyAutomaticRecovery(email, recovery);
            automaticAttachmentWarning = recovery;
            sendRuntimeDiagnostic("assets-start", sendStartedAt, { outcome: "started" });
            addAssetsForSend(
              Office.context.mailbox.item,
              email,
              null,
              function () {
                sendRuntimeDiagnostic(
                  "assets-complete",
                  sendStartedAt,
                  getSendAssetDiagnostics(email)
                );
                if (email.includeEmailAttachments) {
                  automaticAttachmentWarning = {
                    status: "attachments-not-ready",
                    ticketId: recovery.ticketId,
                    ticketNumber: recovery.ticketNumber,
                  };
                }
                sendRuntimeDiagnostic("request-start", sendStartedAt, { outcome: "started" });
                sendRequest(SEND_AUTO_ATTACH_URL, email, function (sendError, result, statusCode) {
                  sendRuntimeDiagnostic("request-complete", sendStartedAt, {
                    outcome: getSendRequestOutcome(sendError, result),
                    statusCode: statusCode,
                  });
                  if (sendError || !result) {
                    sendRuntimeDiagnostic("event-complete", sendStartedAt, {
                      outcome: "blocked-auto-request-failed",
                    });
                    completeWithHaloWarning(complete, automaticAttachmentWarning, sendError);
                    return;
                  }

                  if (
                    result.status === "attached" ||
                    result.status === "attached-with-image-warnings" ||
                    result.status === "attached-with-attachment-warnings" ||
                    result.status === "attached-with-warnings" ||
                    result.status === "already-attached"
                  ) {
                    sendRuntimeDiagnostic("event-complete", sendStartedAt, {
                      outcome: "allowed-auto-result",
                    });
                    completeAllow(complete);
                    return;
                  }

                  sendRuntimeDiagnostic("event-complete", sendStartedAt, { outcome: "blocked" });
                  completeWithHaloWarning(complete, result || automaticAttachmentWarning);
                });
              },
              sendStartedAt
            );
          });
        });
      });
    } catch (error) {
      void error;
      sendRuntimeDiagnostic("event-complete", sendStartedAt, { outcome: "allowed-exception" });
      completeAllow(complete);
    }
  }

  function onHaloMessageAttachmentsChanged(event) {
    var item = Office.context.mailbox.item;
    var eventStartedAt = Date.now();
    var finished = false;

    function complete() {
      if (finished) {
        return;
      }
      finished = true;
      event.completed();
    }

    try {
      sendRuntimeDiagnostic("attachment-change-event", eventStartedAt, { outcome: "started" });
      readComposeAttachMarker(function (markerError, marker) {
        if (markerError) {
          sendRuntimeDiagnostic("attachment-change-event", eventStartedAt, { outcome: "failed" });
          complete();
          return;
        }
        saveComposeItem(item, function (saveError) {
          if (saveError) {
            sendRuntimeDiagnostic("draft-save-complete", eventStartedAt, { outcome: "failed" });
            complete();
            return;
          }
          readCurrentComposeEmail(Boolean(marker), function (readError, email) {
            if (readError || !email) {
              sendRuntimeDiagnostic("compose-read-complete", eventStartedAt, {
                outcome: readError ? "failed" : "allowed-no-email",
              });
              complete();
              return;
            }
            applyComposeMarkerRuntimeContext(email, marker);
            collectOrdinaryAttachmentMetadata(
              item,
              email.bodyHtml,
              function (metadataError, metadata) {
                if (metadataError || !metadata) {
                  sendRuntimeDiagnostic("attachment-inventory", eventStartedAt, {
                    outcome: "inventory-read-failed",
                  });
                  complete();
                  return;
                }
                sendRuntimeDiagnostic("attachment-inventory", eventStartedAt, {
                  attachmentCount: metadata.attachments.length,
                  outcome: metadata.attachments.length ? "ok" : "no-attachments",
                });
                readComposeEmailAttachmentState(item, function (previousState) {
                  if (!metadata.attachments.length) {
                    if (previousState && previousState.emailAttachmentPrefetchKey) {
                      cancelEmailAttachmentPrefetch(
                        previousState.emailAttachmentPrefetchKey,
                        email.backgroundSessionId
                      );
                    }
                    if (marker) {
                      marker.emailAttachmentFingerprint = "";
                      marker.emailAttachmentPrefetchKey = "";
                      marker.emailAttachmentSummary = createEmailAttachmentSummary(0);
                      saveComposeAttachMarker(item, marker, function () {
                        sendRuntimeDiagnostic("attachment-change-event", eventStartedAt, {
                          outcome: "no-attachments",
                        });
                        complete();
                      });
                      return;
                    }
                    sendRuntimeDiagnostic("attachment-change-event", eventStartedAt, {
                      outcome: "no-attachments",
                    });
                    complete();
                    return;
                  }

                  // An event-based handler cannot ask a custom question. Preserve the
                  // explicit task-pane choice when the user selected Email only.
                  if (marker && marker.emailAttachmentDecision === "exclude") {
                    marker.emailAttachmentFingerprint = metadata.fingerprint;
                    marker.emailAttachmentPrefetchKey = "";
                    marker.emailAttachmentSummary = createEmailAttachmentSummary(
                      metadata.attachments.length
                    );
                    if (previousState && previousState.emailAttachmentPrefetchKey) {
                      cancelEmailAttachmentPrefetch(
                        previousState.emailAttachmentPrefetchKey,
                        email.backgroundSessionId
                      );
                    }
                    saveComposeAttachMarker(item, marker, function () {
                      sendRuntimeDiagnostic("attachment-state", eventStartedAt, {
                        attachmentCount: metadata.attachments.length,
                        includeAttachments: false,
                        outcome: "excluded",
                      });
                      complete();
                    });
                    return;
                  }
                  if (
                    !marker &&
                    previousState &&
                    previousState.emailAttachmentDecision === "exclude"
                  ) {
                    previousState.emailAttachmentFingerprint = metadata.fingerprint;
                    previousState.emailAttachmentPrefetchKey = "";
                    previousState.emailAttachmentSummary = summarizeEligibleAttachmentMetadata(
                      metadata.attachments
                    );
                    saveComposeEmailAttachmentState(item, previousState, function () {
                      sendRuntimeDiagnostic("attachment-state", eventStartedAt, {
                        attachmentCount: metadata.attachments.length,
                        includeAttachments: false,
                        outcome: "excluded",
                      });
                      complete();
                    });
                    return;
                  }

                  sendRuntimeDiagnostic("attachment-staging", eventStartedAt, {
                    attachmentCount: metadata.attachments.length,
                    outcome: "started",
                  });
                  prefetchEmailAttachments(
                    item,
                    email,
                    marker,
                    metadata,
                    previousState,
                    Date.now() + EMAIL_ATTACHMENT_EVENT_BUDGET_MS,
                    function (preparation) {
                      var state;
                      if (!preparation || !preparation.prefetchKey) {
                        sendRuntimeDiagnostic("attachment-staging", eventStartedAt, {
                          attachmentCount: metadata.attachments.length,
                          failedCount:
                            Number(
                              preparation && preparation.summary && preparation.summary.failed
                            ) || 0,
                          outcome:
                            (preparation && preparation.diagnosticOutcome) || "stage-missing",
                        });
                        complete();
                        return;
                      }
                      sendRuntimeDiagnostic("attachment-staging", eventStartedAt, {
                        attachmentCount:
                          Number(preparation.summary && preparation.summary.selected) || 0,
                        failedCount: Number(preparation.summary && preparation.summary.failed) || 0,
                        outcome:
                          preparation.summary && preparation.summary.failed ? "not-ready" : "ready",
                        skippedCount:
                          Number(preparation.summary && preparation.summary.skipped) || 0,
                        uploadedCount:
                          Number(preparation.summary && preparation.summary.prepared) || 0,
                      });
                      state = {
                        version: 2,
                        draftItemId: email.itemId || (marker && marker.draftItemId) || "",
                        ticketId: String(preparation.ticketId || (marker && marker.ticketId) || ""),
                        ticketNumber: String(
                          preparation.ticketNumber || (marker && marker.ticketNumber) || ""
                        ),
                        actionMode: normalizeActionMode(
                          (marker && marker.actionMode) || preparation.actionMode
                        ),
                        operationId: preparation.operationId || createOpaqueOperationId(),
                        emailAttachmentDecision: "include",
                        emailAttachmentFingerprint: preparation.fingerprint,
                        emailAttachmentPrefetchKey: preparation.prefetchKey,
                        emailAttachmentStagingVersion: 2,
                        emailAttachmentSummary: preparation.summary,
                      };
                      saveComposeEmailAttachmentState(item, state, function () {
                        if (!marker) {
                          complete();
                          return;
                        }
                        marker.emailAttachmentFingerprint = preparation.fingerprint;
                        marker.emailAttachmentPrefetchKey = preparation.prefetchKey;
                        marker.emailAttachmentStagingVersion = 2;
                        marker.emailAttachmentSummary = preparation.summary;
                        saveComposeAttachMarker(item, marker, complete);
                      });
                    },
                    function () {
                      return finished;
                    }
                  );
                });
              }
            );
          });
        });
      });
    } catch (error) {
      void error;
      sendRuntimeDiagnostic("attachment-change-event", eventStartedAt, { outcome: "failed" });
      complete();
    }
  }

  function sendExplicitComposeAttachment(marker, complete, sendStartedAt) {
    var item = Office.context.mailbox.item;

    sendRuntimeDiagnostic("draft-save-start", sendStartedAt, { outcome: "started" });
    saveComposeItem(item, function (saveError, savedItemId) {
      if (saveError || !savedItemId || savedItemId !== marker.draftItemId) {
        sendRuntimeDiagnostic("draft-save-complete", sendStartedAt, {
          outcome: getDraftSaveOutcome(saveError),
        });
        completeWithHaloWarning(complete, marker, "draft");
        return;
      }

      sendRuntimeDiagnostic("draft-save-complete", sendStartedAt, { outcome: "ok" });
      sendRuntimeDiagnostic("compose-read-start", sendStartedAt, { outcome: "started" });
      readCurrentComposeEmail(true, function (readError, email) {
        if (readError || !email) {
          sendRuntimeDiagnostic("compose-read-complete", sendStartedAt, { outcome: "failed" });
          completeWithHaloWarning(complete, marker, "compose");
          return;
        }

        applyComposeMarkerRuntimeContext(email, marker);
        sendRuntimeDiagnostic("compose-read-complete", sendStartedAt, {
          hasBackgroundSession: Boolean(email.backgroundSessionId),
          hasMailboxIdentity: Boolean(email.mailboxEmail),
          outcome: "ok",
        });
        email.composeAttachId = marker.composeAttachId;
        email.ticketNumber = marker.ticketNumber;

        sendRuntimeDiagnostic("assets-start", sendStartedAt, { outcome: "started" });
        addAssetsForSend(
          item,
          email,
          marker,
          function () {
            sendRuntimeDiagnostic("assets-complete", sendStartedAt, getSendAssetDiagnostics(email));
            sendRuntimeDiagnostic("request-start", sendStartedAt, { outcome: "started" });
            sendRequest(
              SEND_EXPLICIT_ATTACH_BASE_URL + encodeURIComponent(marker.ticketId) + "/sent-email",
              email,
              function (sendError, result, statusCode) {
                sendRuntimeDiagnostic("request-complete", sendStartedAt, {
                  outcome: getSendRequestOutcome(sendError, result),
                  statusCode: statusCode,
                });
                if (
                  sendError ||
                  !result ||
                  !result.ok ||
                  (result.status !== "attached" &&
                    result.status !== "attached-with-image-warnings" &&
                    result.status !== "attached-with-attachment-warnings" &&
                    result.status !== "attached-with-warnings" &&
                    result.status !== "already-attached")
                ) {
                  if (result) {
                    result.ticketId = result.ticketId || marker.ticketId;
                    result.ticketNumber = result.ticketNumber || marker.ticketNumber;
                  }
                  sendRuntimeDiagnostic("event-complete", sendStartedAt, { outcome: "blocked" });
                  completeWithHaloWarning(complete, result || marker, sendError);
                  return;
                }

                sendRuntimeDiagnostic("event-complete", sendStartedAt, { outcome: "allowed" });
                completeAllow(complete);
              }
            );
          },
          sendStartedAt
        );
      });
    });
  }

  function sendComposeTicketCreation(marker, complete, sendStartedAt) {
    var item = Office.context.mailbox.item;
    var operationId = marker.creationOperationId || marker.composeAttachId;

    sendRuntimeDiagnostic("draft-save-start", sendStartedAt, { outcome: "started" });
    saveComposeItem(item, function (saveError, savedItemId) {
      if (saveError || !savedItemId || savedItemId !== marker.draftItemId) {
        sendRuntimeDiagnostic("draft-save-complete", sendStartedAt, {
          outcome: getDraftSaveOutcome(saveError),
        });
        completeWithHaloWarning(complete, marker, "draft");
        return;
      }

      sendRuntimeDiagnostic("draft-save-complete", sendStartedAt, { outcome: "ok" });
      sendRuntimeDiagnostic("compose-read-start", sendStartedAt, { outcome: "started" });
      readCurrentComposeEmail(true, function (readError, email) {
        if (readError || !email) {
          sendRuntimeDiagnostic("compose-read-complete", sendStartedAt, { outcome: "failed" });
          completeWithHaloWarning(complete, marker, "compose");
          return;
        }

        applyComposeMarkerRuntimeContext(email, marker);
        sendRuntimeDiagnostic("compose-read-complete", sendStartedAt, {
          hasBackgroundSession: Boolean(email.backgroundSessionId),
          hasMailboxIdentity: Boolean(email.mailboxEmail),
          outcome: "ok",
        });
        email.composeAttachId = operationId;
        sendRuntimeDiagnostic("assets-start", sendStartedAt, { outcome: "started" });
        addAssetsForSend(
          item,
          email,
          marker,
          function () {
            sendRuntimeDiagnostic("assets-complete", sendStartedAt, getSendAssetDiagnostics(email));
            sendRuntimeDiagnostic("request-start", sendStartedAt, { outcome: "started" });
            sendRequest(
              SEND_CREATE_TICKET_BASE_URL + encodeURIComponent(operationId) + "/send",
              email,
              function (sendError, result, statusCode) {
                sendRuntimeDiagnostic("request-complete", sendStartedAt, {
                  outcome: getSendRequestOutcome(sendError, result),
                  statusCode: statusCode,
                });
                if (
                  sendError ||
                  !result ||
                  !result.ok ||
                  (result.status !== "attached" &&
                    result.status !== "attached-with-image-warnings" &&
                    result.status !== "attached-with-attachment-warnings" &&
                    result.status !== "attached-with-warnings")
                ) {
                  if (result) {
                    result.version = marker.version;
                    result.destinationKind = marker.destinationKind;
                    result.ticketTypeName = marker.ticketTypeName;
                    result.ticketNumber = result.ticketNumber || marker.ticketNumber;
                  }
                  sendRuntimeDiagnostic("event-complete", sendStartedAt, { outcome: "blocked" });
                  completeWithHaloWarning(complete, result || marker, sendError);
                  return;
                }
                sendRuntimeDiagnostic("event-complete", sendStartedAt, { outcome: "allowed" });
                completeAllow(complete);
              }
            );
          },
          sendStartedAt
        );
      });
    });
  }

  function readComposeAttachMarker(callback) {
    var item = Office.context.mailbox.item;

    if (!item || !isMessageItem(item)) {
      callback(null, null);
      return;
    }

    readSessionComposeAttachMarker(item, function (sessionMarker) {
      if (sessionMarker) {
        validateComposeAttachMarker(item, sessionMarker, callback);
        return;
      }

      readCustomComposeAttachMarker(item, function (customMarker) {
        if (!customMarker) {
          callback(null, null);
          return;
        }

        validateComposeAttachMarker(item, customMarker, callback);
      });
    });
  }

  function readSessionComposeAttachMarker(item, callback) {
    if (!item.sessionData || !item.sessionData.getAsync) {
      callback(null);
      return;
    }

    item.sessionData.getAsync(COMPOSE_ATTACH_STORAGE_KEY, function (result) {
      if (!isSucceeded(result)) {
        callback(null);
        return;
      }

      callback(parseComposeAttachMarker(result.value));
    });
  }

  function readCustomComposeAttachMarker(item, callback) {
    if (!item.loadCustomPropertiesAsync) {
      callback(null);
      return;
    }

    item.loadCustomPropertiesAsync(function (result) {
      var value;
      if (!isSucceeded(result) || !result.value || !result.value.get) {
        callback(null);
        return;
      }

      value = result.value.get(COMPOSE_ATTACH_STORAGE_KEY);
      callback(parseComposeAttachMarker(value));
    });
  }

  function readComposeEmailAttachmentState(item, callback) {
    if (item && item.sessionData && item.sessionData.getAsync) {
      item.sessionData.getAsync(COMPOSE_EMAIL_ATTACHMENT_STORAGE_KEY, function (result) {
        var state = null;
        if (isSucceeded(result)) {
          state = parseComposeEmailAttachmentState(result.value);
        }
        if (state) {
          callback(state);
          return;
        }
        readCustomComposeEmailAttachmentState(item, callback);
      });
      return;
    }
    readCustomComposeEmailAttachmentState(item, callback);
  }

  function readCustomComposeEmailAttachmentState(item, callback) {
    if (!item || !item.loadCustomPropertiesAsync) {
      callback(null);
      return;
    }
    item.loadCustomPropertiesAsync(function (result) {
      if (!isSucceeded(result) || !result.value || !result.value.get) {
        callback(null);
        return;
      }
      callback(
        parseComposeEmailAttachmentState(result.value.get(COMPOSE_EMAIL_ATTACHMENT_STORAGE_KEY))
      );
    });
  }

  function parseComposeEmailAttachmentState(value) {
    var state;
    if (typeof value !== "string" || !value) {
      return null;
    }
    try {
      state = JSON.parse(value);
    } catch (error) {
      void error;
      return null;
    }
    if (
      !state ||
      state.version !== 2 ||
      !state.draftItemId ||
      !state.operationId ||
      !state.emailAttachmentFingerprint ||
      (state.emailAttachmentDecision !== "include" &&
        state.emailAttachmentDecision !== "exclude") ||
      (state.emailAttachmentDecision === "include" && !state.emailAttachmentPrefetchKey)
    ) {
      return null;
    }
    return state;
  }

  function saveComposeEmailAttachmentState(item, state, callback) {
    var storedState = copyComposeAttachMarker(state);
    storedState.emailAttachmentSummary = attachmentSummaryForOutlookStorage(
      state.emailAttachmentSummary
    );
    var value = JSON.stringify(storedState);
    var saveSession = Boolean(item && item.sessionData && item.sessionData.setAsync);
    var saveCustom = Boolean(item && item.loadCustomPropertiesAsync);
    var remaining = 0;
    if (saveSession) {
      remaining += 1;
    }
    if (saveCustom) {
      remaining += 1;
    }
    var completed = false;
    function done() {
      if (completed) {
        return;
      }
      remaining -= 1;
      if (remaining <= 0) {
        completed = true;
        callback();
      }
    }
    if (!remaining) {
      callback();
      return;
    }
    if (saveSession) {
      item.sessionData.setAsync(COMPOSE_EMAIL_ATTACHMENT_STORAGE_KEY, value, done);
    }
    if (saveCustom) {
      item.loadCustomPropertiesAsync(function (result) {
        if (!isSucceeded(result) || !result.value || !result.value.set) {
          done();
          return;
        }
        result.value.set(COMPOSE_EMAIL_ATTACHMENT_STORAGE_KEY, value);
        result.value.saveAsync(done);
      });
    }
  }

  function saveComposeAttachMarker(item, marker, callback) {
    var customMarker = copyComposeAttachMarker(marker);
    var sessionMarker = copyComposeAttachMarker(marker);
    sessionMarker.emailAttachmentSummary = attachmentSummaryForOutlookStorage(
      marker.emailAttachmentSummary
    );
    customMarker.emailAttachmentSummary = attachmentSummaryForOutlookStorage(
      marker.emailAttachmentSummary
    );
    var sessionValue = JSON.stringify(sessionMarker);
    var customValue;
    delete customMarker.backgroundSessionId;
    customValue = JSON.stringify(customMarker);
    var saveSession = Boolean(item && item.sessionData && item.sessionData.setAsync);
    var saveCustom = Boolean(item && item.loadCustomPropertiesAsync);
    var remaining = 0;
    if (saveSession) {
      remaining += 1;
    }
    if (saveCustom) {
      remaining += 1;
    }
    var completed = false;
    function done() {
      if (completed) {
        return;
      }
      remaining -= 1;
      if (remaining <= 0) {
        completed = true;
        saveComposeItem(item, function () {
          callback();
        });
      }
    }
    if (!remaining) {
      callback();
      return;
    }
    if (saveSession) {
      item.sessionData.setAsync(COMPOSE_ATTACH_STORAGE_KEY, sessionValue, done);
    }
    if (saveCustom) {
      item.loadCustomPropertiesAsync(function (result) {
        if (!isSucceeded(result) || !result.value || !result.value.set) {
          done();
          return;
        }
        result.value.set(COMPOSE_ATTACH_STORAGE_KEY, customValue);
        result.value.saveAsync(done);
      });
    }
  }

  function copyComposeAttachMarker(marker) {
    var copy = {};
    var key;
    for (key in marker) {
      if (Object.prototype.hasOwnProperty.call(marker, key)) {
        copy[key] = marker[key];
      }
    }
    return copy;
  }

  function attachmentSummaryForOutlookStorage(summary) {
    var stored;
    if (!summary) {
      return summary;
    }
    stored = copyComposeAttachMarker(summary);
    stored.warnings = [];
    return stored;
  }

  function applyComposeMarkerRuntimeContext(email, marker) {
    if (!email || !marker) {
      return;
    }
    if (marker.backgroundSessionId) {
      email.backgroundSessionId = marker.backgroundSessionId;
    }
    if (marker.mailboxEmail) {
      email.mailboxEmail = marker.mailboxEmail;
    }
    email.actionMode = marker.actionMode === "private-note" ? "private-note" : "email";
  }

  function validateComposeAttachMarker(item, marker, callback) {
    readItemId(item, function (itemId) {
      if (!itemId || itemId !== marker.draftItemId) {
        callback(null, null);
        return;
      }

      callback(null, marker);
    });
  }

  function parseComposeAttachMarker(value) {
    var invalidDestination;
    var marker;
    if (typeof value !== "string" || !value) {
      return null;
    }

    try {
      marker = JSON.parse(value);
    } catch (error) {
      void error;
      return null;
    }

    if (!marker || typeof marker !== "object") {
      return null;
    }

    if (isCreateTicketMarker(marker)) {
      invalidDestination = !marker.creationOperationId || !marker.ticketTypeId;
    } else {
      invalidDestination = !marker.ticketId || !marker.ticketNumber;
    }

    if (
      !marker ||
      (marker.version !== 1 &&
        marker.version !== 2 &&
        marker.version !== 3 &&
        marker.version !== 4) ||
      !marker.composeAttachId ||
      !marker.draftItemId ||
      invalidDestination
    ) {
      return null;
    }

    return marker;
  }

  function isCreateTicketMarker(marker) {
    return Boolean(marker && marker.version === 4 && marker.destinationKind === "create-ticket");
  }

  function normalizeActionMode(value) {
    return value === "private-note" ? "private-note" : "email";
  }

  function saveComposeItem(item, callback) {
    if (!item || !item.saveAsync) {
      callback(new Error("Outlook could not save this draft."), "");
      return;
    }

    item.saveAsync(function (result) {
      if (isSucceeded(result) && result.value) {
        callback(null, result.value);
        return;
      }

      callback(new Error(getAsyncErrorMessage(result, "Outlook could not save this draft.")), "");
    });
  }

  function readCurrentComposeEmail(allowUnmapped, callback) {
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
    readMessageBody(item, function (bodyError, body) {
      if (bodyError || !body || (!body.bodyHtml && !body.bodyText)) {
        callback(null, null);
        return;
      }

      if (!allowUnmapped && !inReplyToMessageIds.length && !conversationId) {
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

  function addAssetsForSend(item, email, marker, callback, diagnosticStartedAt) {
    var remaining = 2;
    function finished() {
      remaining -= 1;
      if (!remaining) {
        callback();
      }
    }
    addInlineImagesForSend(item, email, marker, finished);
    addEmailAttachmentsForSend(item, email, marker, finished, diagnosticStartedAt);
  }

  function addEmailAttachmentsForSend(item, email, marker, callback, diagnosticStartedAt) {
    var startedAt = diagnosticStartedAt || Date.now();
    collectOrdinaryAttachmentMetadata(item, email.bodyHtml, function (metadataError, metadata) {
      var inventory;
      if (metadataError || !metadata) {
        sendRuntimeDiagnostic("attachment-inventory", startedAt, {
          outcome: "inventory-read-failed",
        });
        callback();
        return;
      }
      if (!metadata.attachments.length) {
        email.includeEmailAttachments = false;
        email.emailAttachmentFingerprint = "";
        email.emailAttachmentSummary = createEmailAttachmentSummary(0);
        sendRuntimeDiagnostic("attachment-state", startedAt, {
          attachmentCount: 0,
          includeAttachments: false,
          outcome: "no-attachments",
        });
        callback();
        return;
      }
      inventory = summarizeEligibleAttachmentMetadata(metadata.attachments);
      if (marker && marker.emailAttachmentDecision === "exclude") {
        email.includeEmailAttachments = false;
        email.emailAttachmentFingerprint = metadata.fingerprint;
        email.emailAttachmentSummary = inventory;
        sendRuntimeDiagnostic("attachment-state", startedAt, {
          attachmentCount: inventory.selected,
          includeAttachments: false,
          outcome: "excluded",
          skippedCount: inventory.skipped,
        });
        callback();
        return;
      }
      readComposeEmailAttachmentState(item, function (state) {
        var expectedTicketId;
        var stateDiagnosticOutcome = "state-missing";
        var stateBelongsToDraft;
        expectedTicketId = marker && marker.ticketId;
        if (isCreateTicketMarker(marker)) {
          expectedTicketId = "0";
        }
        stateBelongsToDraft = Boolean(
          state &&
          state.version === 2 &&
          state.draftItemId === email.itemId &&
          (!marker || state.ticketId === String(expectedTicketId)) &&
          (!marker || state.operationId === marker.composeAttachId)
        );
        if (!marker && stateBelongsToDraft && state.emailAttachmentDecision === "exclude") {
          email.includeEmailAttachments = false;
          email.emailAttachmentFingerprint = metadata.fingerprint;
          email.emailAttachmentSummary = inventory;
          sendRuntimeDiagnostic("attachment-state", startedAt, {
            attachmentCount: inventory.selected,
            includeAttachments: false,
            outcome: "excluded",
            skippedCount: inventory.skipped,
          });
          callback();
          return;
        }
        if (
          stateBelongsToDraft &&
          state.emailAttachmentStagingVersion === 2 &&
          state.emailAttachmentPrefetchKey &&
          state.emailAttachmentFingerprint === metadata.fingerprint
        ) {
          inventory.selected =
            Number(state.emailAttachmentSummary && state.emailAttachmentSummary.selected) || 0;
          inventory.prepared =
            Number(state.emailAttachmentSummary && state.emailAttachmentSummary.prepared) || 0;
          inventory.failed =
            Number(state.emailAttachmentSummary && state.emailAttachmentSummary.failed) || 0;
          inventory.skipped =
            Number(state.emailAttachmentSummary && state.emailAttachmentSummary.skipped) || 0;
          applyEmailAttachmentPreparation(email, {
            actionMode: state.actionMode,
            draftItemId: state.draftItemId,
            fingerprint: metadata.fingerprint,
            operationId: state.operationId,
            prefetchKey: state.emailAttachmentPrefetchKey,
            stagingVersion: 2,
            summary: inventory,
            ticketNumber: state.ticketNumber,
          });
          sendRuntimeDiagnostic("attachment-state", startedAt, {
            attachmentCount: inventory.selected,
            failedCount: inventory.failed,
            includeAttachments: true,
            outcome: "prepared",
            skippedCount: inventory.skipped,
            uploadedCount: inventory.prepared,
          });
          callback();
          return;
        }
        if (state) {
          if (!stateBelongsToDraft) {
            stateDiagnosticOutcome = "state-mismatch";
          } else if (
            state.emailAttachmentFingerprint &&
            state.emailAttachmentFingerprint !== metadata.fingerprint
          ) {
            stateDiagnosticOutcome = "inventory-mismatch";
          } else {
            stateDiagnosticOutcome = "stage-missing";
          }
        }
        email.includeEmailAttachments = true;
        email.emailAttachmentDraftItemId = email.itemId || "";
        email.emailAttachmentFingerprint = metadata.fingerprint;
        email.emailAttachmentOperationId = marker
          ? marker.composeAttachId
          : stateBelongsToDraft
            ? state.operationId
            : "";
        email.emailAttachmentPrefetchKey = "";
        email.emailAttachmentStagingVersion = 2;
        email.emailAttachmentSummary = inventory;
        sendRuntimeDiagnostic("attachment-state", startedAt, {
          attachmentCount: inventory.selected,
          failedCount: inventory.failed,
          includeAttachments: true,
          outcome: stateDiagnosticOutcome,
          skippedCount: inventory.skipped,
          uploadedCount: inventory.prepared,
        });
        callback();
      });
    });
  }

  function applyEmailAttachmentPreparation(email, preparation) {
    email.includeEmailAttachments = true;
    email.emailAttachmentDraftItemId = preparation.draftItemId || email.itemId || "";
    email.emailAttachmentFingerprint = preparation.fingerprint || "";
    email.emailAttachmentOperationId = preparation.operationId || "";
    email.emailAttachmentPrefetchKey = preparation.prefetchKey || "";
    email.emailAttachmentStagingVersion = preparation.stagingVersion || 2;
    email.emailAttachmentSummary = preparation.summary;
    if (preparation.actionMode) {
      email.actionMode = normalizeActionMode(preparation.actionMode);
    }
    if (preparation.ticketNumber) {
      email.mappedTicketNumber = String(preparation.ticketNumber);
    }
  }

  function summarizeEligibleAttachmentMetadata(attachments) {
    var summary = createEmailAttachmentSummary(attachments.length);
    var totalBytes = 0;
    var index;
    for (index = 0; index < attachments.length; index += 1) {
      if (
        index >= MAX_EMAIL_ATTACHMENTS ||
        attachments[index].reportedSize > MAX_EMAIL_ATTACHMENT_BYTES ||
        isUnsupportedAttachmentType(attachments[index].attachmentType) ||
        totalBytes + attachments[index].reportedSize > MAX_EMAIL_ATTACHMENT_TOTAL_BYTES
      ) {
        summary.skipped += 1;
        continue;
      }
      totalBytes += attachments[index].reportedSize;
      summary.selected += 1;
    }
    return summary;
  }

  function isUnsupportedAttachmentType(value) {
    return ["cloud", "link", "reference", "url"].indexOf(String(value || "").toLowerCase()) >= 0;
  }

  function collectOrdinaryAttachmentMetadata(item, bodyHtml, callback) {
    readRawAttachmentMetadata(item, function (values) {
      var contentIds = extractCidReferences(bodyHtml);
      var referenced = {};
      var seenAttachmentIds = {};
      var attachments = [];
      var index;
      var value;
      var contentId;
      var attachmentType;
      for (index = 0; index < contentIds.length; index += 1) {
        referenced[contentIds[index]] = true;
      }
      for (index = 0; index < values.length; index += 1) {
        value = values[index] || {};
        if (!value.id || seenAttachmentIds[String(value.id)]) {
          continue;
        }
        seenAttachmentIds[String(value.id)] = true;
        contentId = normalizeContentId(value.contentId);
        if (
          value.isInline === true ||
          (contentId && referenced[contentId]) ||
          referenced[normalizeContentId(value.id)]
        ) {
          continue;
        }
        attachmentType = String(value.attachmentType || "file").toLowerCase();
        attachments.push({
          attachmentType: attachmentType,
          contentId: contentId,
          contentType: String(value.contentType || "application/octet-stream").toLowerCase(),
          id: String(value.id),
          name: sanitizeEmailAttachmentName(String(value.name || "email-attachment.bin")),
          rawName: String(value.name || "email-attachment.bin"),
          reportedSize: Math.max(0, Number(value.size) || 0),
        });
      }
      hashOrdinaryAttachmentMetadata(attachments, callback);
    });
  }

  function readRawAttachmentMetadata(item, callback) {
    var fallbackAttachments = [];
    if (!item) {
      callback([]);
      return;
    }
    if (item.attachments && typeof item.attachments.length === "number") {
      fallbackAttachments = item.attachments;
    }
    if (item.getAttachmentsAsync) {
      item.getAttachmentsAsync(function (result) {
        if (isSucceeded(result)) {
          callback(result.value || []);
          return;
        }
        callback(fallbackAttachments);
      });
      return;
    }
    callback(fallbackAttachments);
  }

  function hashOrdinaryAttachmentMetadata(attachments, callback) {
    var remaining;
    var failed = false;
    var index;
    var keyedAttachments;
    var descriptorCounts = {};
    var fingerprintSources = [];
    if (!attachments.length) {
      callback(null, { attachments: [], fingerprint: "" });
      return;
    }
    keyedAttachments = attachments.map(function (attachment, originalIndex) {
      return {
        attachment: attachment,
        descriptor: [
          String(attachment.name || "").toLowerCase(),
          attachment.reportedSize,
          String(attachment.contentType || "").toLowerCase(),
          String(attachment.attachmentType || "").toLowerCase(),
        ].join("\u0000"),
        originalIndex: originalIndex,
      };
    });
    keyedAttachments.sort(function (left, right) {
      if (left.descriptor === right.descriptor) {
        return left.originalIndex - right.originalIndex;
      }
      return left.descriptor < right.descriptor ? -1 : 1;
    });
    for (index = 0; index < keyedAttachments.length; index += 1) {
      attachments[index] = keyedAttachments[index].attachment;
    }
    for (index = 0; index < keyedAttachments.length; index += 1) {
      descriptorCounts[keyedAttachments[index].descriptor] =
        (descriptorCounts[keyedAttachments[index].descriptor] || 0) + 1;
      keyedAttachments[index].duplicateIndex = descriptorCounts[keyedAttachments[index].descriptor];
    }
    for (var descriptor in descriptorCounts) {
      if (Object.prototype.hasOwnProperty.call(descriptorCounts, descriptor)) {
        fingerprintSources.push(descriptor + "\u0000" + descriptorCounts[descriptor]);
      }
    }
    fingerprintSources.sort();
    remaining = keyedAttachments.length + 1;
    function complete(error) {
      if (failed) {
        return;
      }
      if (error) {
        failed = true;
        callback(error);
        return;
      }
      remaining -= 1;
      if (remaining) {
        return;
      }
      callback(null, { attachments: attachments, fingerprint: attachments.fingerprint || "" });
    }
    for (index = 0; index < keyedAttachments.length; index += 1) {
      (function (entry) {
        sha256Bytes(
          utf8Bytes(entry.descriptor + "\u0000" + entry.duplicateIndex),
          function (error, hash) {
            if (!error) {
              entry.attachment.attachmentKey = hash;
            }
            complete(error);
          }
        );
      })(keyedAttachments[index]);
    }
    sha256Bytes(utf8Bytes(fingerprintSources.join("\u0001")), function (error, fingerprint) {
      attachments.fingerprint = fingerprint || "";
      complete(error);
    });
  }

  function prefetchEmailAttachments(
    item,
    email,
    marker,
    metadata,
    previousState,
    deadline,
    callback,
    isCancelled
  ) {
    var prefetchStartedAt = Date.now();
    if (!marker && !email.mappingRecoveryValidated) {
      recoverAutomaticMapping(item, email, function (recovery) {
        if (!recovery || recovery.status !== "matched") {
          callback({
            diagnosticOutcome: "no-match",
            fingerprint: metadata.fingerprint,
            prefetchKey: "",
            summary: summarizeEligibleAttachmentMetadata(metadata.attachments),
          });
          return;
        }
        applyAutomaticRecovery(email, recovery);
        prefetchEmailAttachments(
          item,
          email,
          marker,
          metadata,
          previousState,
          deadline,
          callback,
          isCancelled
        );
      });
      return;
    }
    collectOrdinaryAttachmentContents(
      item,
      metadata,
      deadline,
      function (prepared) {
        var operationId;
        var startPayload;
        var remainingMs;
        if (isCancelled()) {
          callback(null);
          return;
        }
        if (!prepared.descriptors.length) {
          sendRuntimeDiagnostic("attachment-prefetch-complete", prefetchStartedAt, {
            attachmentCount: 0,
            failedCount: prepared.summary.failed,
            outcome: prepared.summary.failed ? "not-ready" : "ready",
            skippedCount: prepared.summary.skipped,
            uploadedCount: 0,
          });
          callback({
            diagnosticOutcome: prepared.summary.failed ? "not-ready" : "ready",
            fingerprint: metadata.fingerprint,
            prefetchKey: "",
            summary: prepared.summary,
          });
          return;
        }
        operationId = createOpaqueOperationId();
        if (previousState && previousState.operationId) {
          operationId = previousState.operationId;
        }
        if (marker) {
          operationId = marker.composeAttachId;
        }
        startPayload = {
          backgroundSessionId: email.backgroundSessionId,
          conversationId: email.conversationId,
          emailAttachmentFingerprint: metadata.fingerprint,
          emailAttachments: prepared.descriptors,
          draftItemId: email.itemId || (marker && marker.draftItemId) || "",
          inReplyToMessageIds: email.inReplyToMessageIds,
          internetMessageId: email.internetMessageId,
          mailboxEmail: email.mailboxEmail,
          operationId: operationId,
          referenceMessageIds: email.referenceMessageIds,
          recoveredComposeAttachId: email.recoveredComposeAttachId || "",
        };
        if (marker) {
          startPayload.actionMode = normalizeActionMode(marker.actionMode);
          if (isCreateTicketMarker(marker)) {
            startPayload.creationOperationId = marker.creationOperationId || marker.composeAttachId;
          } else {
            startPayload.ticketId = marker.ticketId;
            startPayload.ticketNumber = marker.ticketNumber;
          }
        }
        remainingMs = 290000;
        if (deadline) {
          remainingMs = Math.max(1, deadline - Date.now());
        }
        sendJsonRequest(
          EMAIL_ATTACHMENT_PREFETCH_START_URL,
          startPayload,
          remainingMs,
          function (startError, startResult) {
            var pending = {};
            var tasks = [];
            var index;
            if (
              isCancelled() ||
              startError ||
              !startResult ||
              !startResult.ok ||
              (startResult.status !== "ready" && startResult.status !== "pending") ||
              !startResult.prefetchKey
            ) {
              prepared.summary.failed = Math.max(
                prepared.summary.failed,
                prepared.descriptors.length
              );
              var failureOutcome = "failed";
              if (startError) {
                failureOutcome = getSendRequestOutcome(startError, null);
              } else if (startResult && startResult.status) {
                failureOutcome = String(startResult.status);
              }
              sendRuntimeDiagnostic("attachment-prefetch-complete", prefetchStartedAt, {
                attachmentCount: prepared.summary.selected,
                failedCount: prepared.summary.failed,
                outcome: failureOutcome,
                skippedCount: prepared.summary.skipped,
                uploadedCount: prepared.summary.prepared,
              });
              callback({
                diagnosticOutcome: failureOutcome,
                fingerprint: metadata.fingerprint,
                prefetchKey: "",
                summary: prepared.summary,
              });
              return;
            }
            for (index = 0; index < (startResult.pendingAttachmentKeys || []).length; index += 1) {
              pending[startResult.pendingAttachmentKeys[index]] = true;
            }
            prepared.summary.prepared =
              Number(startResult.aggregate && startResult.aggregate.prepared) || 0;
            for (index = 0; index < prepared.uploads.length; index += 1) {
              if (pending[prepared.uploads[index].attachmentKey]) {
                tasks.push(prepared.uploads[index]);
              }
            }
            runWithConcurrency(
              tasks,
              3,
              function (upload, next) {
                var uploadTimeout = 290000;
                if (deadline) {
                  uploadTimeout = Math.max(1, deadline - Date.now());
                }
                if (isCancelled() || (deadline && Date.now() >= deadline)) {
                  prepared.summary.failed += 1;
                  next();
                  return;
                }
                sendJsonRequest(
                  EMAIL_ATTACHMENT_PREFETCH_BASE_URL +
                    encodeURIComponent(startResult.prefetchKey) +
                    "/items",
                  {
                    attachmentKey: upload.attachmentKey,
                    backgroundSessionId: email.backgroundSessionId,
                    contentBase64: upload.contentBase64,
                    contentFormat: upload.contentFormat,
                    contentSha256: upload.contentSha256,
                  },
                  uploadTimeout,
                  function (uploadError, uploadResult) {
                    if (
                      !uploadError &&
                      uploadResult &&
                      uploadResult.ok &&
                      (uploadResult.status === "prepared" ||
                        uploadResult.status === "already-prepared")
                    ) {
                      prepared.summary.prepared += 1;
                    } else {
                      prepared.summary.failed += 1;
                    }
                    next();
                  }
                );
              },
              function () {
                var preparedTicketId = marker && marker.ticketId;
                var pendingCount = (startResult.pendingAttachmentKeys || []).length;
                if (isCreateTicketMarker(marker)) {
                  preparedTicketId = "0";
                }
                prepared.summary.failed = Math.max(
                  0,
                  pendingCount -
                    Math.max(
                      0,
                      prepared.summary.prepared -
                        (Number(startResult.aggregate && startResult.aggregate.prepared) || 0)
                    )
                );
                sendRuntimeDiagnostic("attachment-prefetch-complete", prefetchStartedAt, {
                  attachmentCount: prepared.summary.selected,
                  failedCount: prepared.summary.failed,
                  outcome: prepared.summary.failed ? "not-ready" : "ready",
                  skippedCount: prepared.summary.skipped,
                  uploadedCount: prepared.summary.prepared,
                });
                callback({
                  actionMode: normalizeActionMode(startResult.actionMode),
                  diagnosticOutcome: prepared.summary.failed ? "not-ready" : "ready",
                  fingerprint: metadata.fingerprint,
                  operationId: operationId,
                  prefetchKey: startResult.prefetchKey,
                  stagingVersion: 2,
                  summary: prepared.summary,
                  ticketId: String(startResult.ticketId || preparedTicketId || ""),
                  ticketNumber: String(startResult.ticketNumber || ""),
                });
              }
            );
          }
        );
      },
      isCancelled
    );
  }

  function collectOrdinaryAttachmentContents(item, metadata, deadline, callback, isCancelled) {
    var readStartedAt = Date.now();
    var candidates = [];
    var summary = createEmailAttachmentSummary(metadata.attachments.length);
    var reportedEligibleBytes = 0;
    var index;
    for (index = 0; index < metadata.attachments.length; index += 1) {
      if (index >= MAX_EMAIL_ATTACHMENTS) {
        summary.skipped += 1;
      } else if (metadata.attachments[index].reportedSize > MAX_EMAIL_ATTACHMENT_BYTES) {
        summary.skipped += 1;
      } else if (isUnsupportedAttachmentType(metadata.attachments[index].attachmentType)) {
        summary.skipped += 1;
      } else if (
        reportedEligibleBytes + metadata.attachments[index].reportedSize >
        MAX_EMAIL_ATTACHMENT_TOTAL_BYTES
      ) {
        summary.skipped += 1;
      } else {
        reportedEligibleBytes += metadata.attachments[index].reportedSize;
        candidates.push({ attachment: metadata.attachments[index], index: index });
      }
    }
    runWithConcurrency(
      candidates,
      3,
      function (task, next) {
        readOrdinaryAttachmentContentWithRetry(
          item,
          task.attachment,
          deadline,
          isCancelled,
          function (readError, value, failureCode, attemptCount) {
            var encoded;
            task.attemptCount = attemptCount;
            if (failureCode && failureCode !== "none") {
              task.failureCode = failureCode;
            }
            if (readError || !value) {
              summary.failed += 1;
              task.failed = true;
              task.failureCode = failureCode || "read-failed";
              next();
              return;
            }
            try {
              encoded = encodeOrdinaryAttachmentContent(
                value.content,
                value.format,
                task.attachment
              );
              if (encoded.unsupported) {
                summary.skipped += 1;
                task.encoded = encoded;
                next();
                return;
              }
              sha256Base64(encoded.contentBase64, function (hashError, contentSha256) {
                if (hashError || !contentSha256) {
                  summary.failed += 1;
                  task.failed = true;
                  task.failureCode = "hash-failed";
                } else {
                  encoded.contentSha256 = contentSha256;
                  task.encoded = encoded;
                }
                next();
              });
              return;
            } catch (error) {
              void error;
              summary.failed += 1;
              task.failed = true;
              task.failureCode = "invalid-content";
            }
            next();
          }
        );
      },
      function () {
        var uploads = [];
        var descriptors = [];
        var totalBytes = 0;
        var attemptCount = 0;
        var failureCode = "none";
        var diagnosticOutcome = "ok";
        for (var resultIndex = 0; resultIndex < candidates.length; resultIndex += 1) {
          var entry = candidates[resultIndex];
          attemptCount += Number(entry.attemptCount) || 0;
          if (failureCode === "none" && entry.failureCode) {
            failureCode = entry.failureCode;
          }
          if (!entry.encoded || entry.encoded.unsupported) {
            if (entry.failed) {
              descriptors.push({
                attachmentKey: entry.attachment.attachmentKey,
                attachmentType: entry.attachment.attachmentType,
                contentSha256: "",
                contentType: entry.attachment.contentType,
                name: entry.attachment.name,
                reportedSize: entry.attachment.reportedSize,
              });
            }
            continue;
          }
          if (entry.encoded.decodedSize > MAX_EMAIL_ATTACHMENT_BYTES) {
            summary.skipped += 1;
            continue;
          }
          if (totalBytes + entry.encoded.decodedSize > MAX_EMAIL_ATTACHMENT_TOTAL_BYTES) {
            summary.skipped += 1;
            continue;
          }
          totalBytes += entry.encoded.decodedSize;
          uploads.push({
            attachmentKey: entry.attachment.attachmentKey,
            contentBase64: entry.encoded.contentBase64,
            contentFormat: entry.encoded.contentFormat,
            contentSha256: entry.encoded.contentSha256,
          });
          descriptors.push({
            attachmentKey: entry.attachment.attachmentKey,
            attachmentType: entry.attachment.attachmentType,
            contentSha256: entry.encoded.contentSha256,
            contentType: entry.encoded.contentType,
            name: entry.encoded.name,
            reportedSize: entry.encoded.decodedSize,
          });
        }
        summary.selected = descriptors.length;
        if (summary.failed) {
          diagnosticOutcome = "failed";
        }
        sendRuntimeDiagnostic("attachment-read-complete", readStartedAt, {
          attachmentCount: metadata.attachments.length,
          attachmentError: failureCode,
          attemptCount: attemptCount,
          failedCount: summary.failed,
          outcome: diagnosticOutcome,
          skippedCount: summary.skipped,
        });
        callback({ descriptors: descriptors, summary: summary, uploads: uploads });
      }
    );
  }

  function readOrdinaryAttachmentContentWithRetry(
    item,
    attachment,
    deadline,
    isCancelled,
    callback
  ) {
    var attemptCount = 0;
    var attachmentId = attachment.id;
    var lastFailureCode = "read-failed";

    function read() {
      var retryDelay;
      var readTimeout;
      var attemptFinished = false;
      var attemptTimeoutMs;

      function failAttempt(failureCode) {
        if (attemptFinished) {
          return;
        }
        attemptFinished = true;
        clearTimeout(readTimeout);
        lastFailureCode = failureCode;
        if (attemptCount > EMAIL_ATTACHMENT_READ_RETRY_DELAYS_MS.length) {
          callback(true, null, lastFailureCode, attemptCount);
          return;
        }
        retryDelay = EMAIL_ATTACHMENT_READ_RETRY_DELAYS_MS[attemptCount - 1];
        if (deadline && Date.now() + retryDelay >= deadline) {
          callback(true, null, "timeout", attemptCount);
          return;
        }
        setTimeout(function () {
          readRawAttachmentMetadata(item, function (values) {
            attachmentId = findCurrentAttachmentId(values, attachment, attachmentId);
            read();
          });
        }, retryDelay);
      }

      if (isCancelled()) {
        callback(true, null, "timeout", attemptCount);
        return;
      }
      if (deadline && Date.now() >= deadline) {
        callback(true, null, "timeout", attemptCount);
        return;
      }
      if (!item || !item.getAttachmentContentAsync) {
        callback(true, null, "not-supported", attemptCount);
        return;
      }
      attemptCount += 1;
      attemptTimeoutMs = EMAIL_ATTACHMENT_CONTENT_READ_TIMEOUT_MS;
      if (deadline) {
        attemptTimeoutMs = Math.max(1, Math.min(attemptTimeoutMs, deadline - Date.now()));
      }
      readTimeout = setTimeout(function () {
        failAttempt("timeout");
      }, attemptTimeoutMs);
      item.getAttachmentContentAsync(attachmentId, function (result) {
        var recoveredFailureCode = "none";
        if (attemptFinished) {
          return;
        }
        attemptFinished = true;
        clearTimeout(readTimeout);
        if (isSucceeded(result) && result.value) {
          if (attemptCount > 1) {
            recoveredFailureCode = lastFailureCode;
          }
          callback(null, result.value, recoveredFailureCode, attemptCount);
          return;
        }
        attemptFinished = false;
        failAttempt(getAttachmentReadFailureCode(result));
      });
    }

    read();
  }

  function findCurrentAttachmentId(values, attachment, fallbackId) {
    var matches = [];
    var index;
    var value;
    for (index = 0; index < values.length; index += 1) {
      value = values[index] || {};
      if (
        value.id &&
        value.isInline !== true &&
        String(value.name || "") === attachment.rawName &&
        Math.max(0, Number(value.size) || 0) === attachment.reportedSize &&
        String(value.contentType || "application/octet-stream").toLowerCase() ===
          attachment.contentType &&
        String(value.attachmentType || "file").toLowerCase() === attachment.attachmentType
      ) {
        matches.push(String(value.id));
      }
    }
    if (matches.length === 1) {
      return matches[0];
    }
    for (index = 0; index < matches.length; index += 1) {
      if (matches[index] === fallbackId) {
        return fallbackId;
      }
    }
    return fallbackId;
  }

  function getAttachmentReadFailureCode(result) {
    var resultError = result && result.error;
    var value = String(
      (resultError && (resultError.code || resultError.message)) || "read-failed"
    ).toLowerCase();
    if (/invalid[\s_-]*attachment[\s_-]*id/.test(value)) {
      return "invalid-attachment-id";
    }
    if (/timed\s*out|time\s*out|timeout/.test(value)) {
      return "timeout";
    }
    if (/not\s*supported|unsupported|not-implemented/.test(value)) {
      return "not-supported";
    }
    return "read-failed";
  }

  function encodeOrdinaryAttachmentContent(content, formatValue, attachment) {
    var format = String(formatValue || "base64").toLowerCase();
    var contentBase64;
    if (format.indexOf("url") >= 0) {
      return { unsupported: true };
    }
    if (format.indexOf("eml") >= 0) {
      contentBase64 = bytesToBase64(utf8Bytes(String(content || "")));
      return {
        contentBase64: contentBase64,
        contentFormat: "eml",
        contentType: "message/rfc822",
        decodedSize: getDecodedBase64Length(contentBase64),
        name: ensureEmailAttachmentExtension(attachment.name, ".eml"),
        unsupported: false,
      };
    }
    if (format.indexOf("icalendar") >= 0 || format.indexOf("calendar") >= 0) {
      contentBase64 = bytesToBase64(utf8Bytes(String(content || "")));
      return {
        contentBase64: contentBase64,
        contentFormat: "icalendar",
        contentType: "text/calendar",
        decodedSize: getDecodedBase64Length(contentBase64),
        name: ensureEmailAttachmentExtension(attachment.name, ".ics"),
        unsupported: false,
      };
    }
    contentBase64 = String(content || "").replace(/\s/g, "");
    if (!contentBase64 || contentBase64.length % 4 || !/^[a-z0-9+/]*={0,2}$/i.test(contentBase64)) {
      throw new Error("Invalid Base64 attachment content.");
    }
    return {
      contentBase64: contentBase64,
      contentFormat: "base64",
      contentType: attachment.contentType || "application/octet-stream",
      decodedSize: getDecodedBase64Length(contentBase64),
      name: attachment.name,
      unsupported: false,
    };
  }

  function bytesToBase64(bytes) {
    var binary = "";
    var chunkSize = 32768;
    var offset;
    var chunk;
    var index;
    for (offset = 0; offset < bytes.length; offset += chunkSize) {
      chunk = bytes.subarray(offset, Math.min(offset + chunkSize, bytes.length));
      for (index = 0; index < chunk.length; index += 1) {
        binary += String.fromCharCode(chunk[index]);
      }
    }
    return btoa(binary);
  }

  function ensureEmailAttachmentExtension(name, extension) {
    var value = sanitizeEmailAttachmentName(name);
    if (value.toLowerCase().slice(-extension.length) === extension) {
      return value;
    }
    return value.slice(0, Math.max(1, 255 - extension.length)) + extension;
  }

  function sanitizeEmailAttachmentName(value) {
    var parts = String(value || "")
      .replace(/\\/g, "/")
      .split("/");
    var name = String(parts.pop() || "")
      .split("")
      .filter(function (character) {
        var code = character.charCodeAt(0);
        return code >= 32 && code !== 127;
      })
      .join("")
      .replace(/^\s+|\s+$/g, "");
    if (!name || name === "." || name === "..") {
      return "email-attachment.bin";
    }
    return name.slice(0, 255);
  }

  function createEmailAttachmentSummary(detected) {
    return {
      attached: 0,
      detected: detected || 0,
      failed: 0,
      prepared: 0,
      selected: 0,
      skipped: 0,
      warnings: [],
    };
  }

  function createOpaqueOperationId() {
    return (
      Date.now().toString(36) +
      "-" +
      Math.random().toString(36).slice(2) +
      "-" +
      Math.random().toString(36).slice(2)
    );
  }

  function addInlineImagesForSend(item, email, marker, callback) {
    var contentIds = extractCidReferences(email.bodyHtml).slice(0, 20);
    var completed = false;
    var timeout;

    if (!contentIds.length) {
      callback();
      return;
    }

    timeout = setTimeout(function () {
      finish();
    }, 2000);

    function finish() {
      if (completed) {
        return;
      }
      completed = true;
      clearTimeout(timeout);
      callback();
    }

    readAttachmentMetadata(item, function (attachments) {
      var source = buildInlineImageFingerprintSource(contentIds, attachments);
      sha256Bytes(utf8Bytes(source), function (fingerprintError, fingerprint) {
        if (completed) {
          return;
        }

        if (
          !fingerprintError &&
          marker &&
          marker.inlineImagePrefetchKey &&
          marker.inlineImageFingerprint === fingerprint
        ) {
          email.inlineImagePrefetchKey = marker.inlineImagePrefetchKey;
          email.inlineImageFingerprint = fingerprint;
          finish();
          return;
        }

        collectInlineImageContents(
          item,
          contentIds,
          attachments,
          function (refs, uploads) {
            if (completed) {
              return;
            }
            email.inlineImageFingerprint = fingerprint;
            if (fingerprintError) {
              email.inlineImageFingerprint = "";
            }
            email.inlineImageRefs = refs;
            email.inlineImageUploads = uploads;
            finish();
          },
          function () {
            return completed;
          }
        );
      });
    });
  }

  function extractCidReferences(html) {
    var contentIds = [];
    var seen = {};
    var pattern = /\b(src|originalsrc)\s*=\s*(["'])(cid:[^"']+)\2/gi;
    var match;
    var contentId;

    while ((match = pattern.exec(String(html || "")))) {
      contentId = normalizeContentId(match[3]);
      if (contentId && !seen[contentId]) {
        seen[contentId] = true;
        contentIds.push(contentId);
      }
    }
    return contentIds;
  }

  function normalizeContentId(value) {
    var normalized = String(value || "")
      .replace(/^\s*cid:/i, "")
      .replace(/&lt;/gi, "<")
      .replace(/&gt;/gi, ">")
      .replace(/&amp;/gi, "&")
      .replace(/^\s+|\s+$/g, "");
    try {
      normalized = decodeURIComponent(normalized);
    } catch (error) {
      void error;
    }
    return normalized
      .replace(/^<+|>+$/g, "")
      .replace(/^\s+|\s+$/g, "")
      .toLowerCase();
  }

  function readAttachmentMetadata(item, callback) {
    if (item && item.attachments && typeof item.attachments.length === "number") {
      callback(normalizeAttachments(item.attachments));
      return;
    }
    if (!item || !item.getAttachmentsAsync) {
      callback([]);
      return;
    }
    item.getAttachmentsAsync(function (result) {
      if (isSucceeded(result)) {
        callback(normalizeAttachments(result.value || []));
        return;
      }
      callback([]);
    });
  }

  function normalizeAttachments(values) {
    var attachments = [];
    var index;
    var value;
    var type;
    for (index = 0; index < values.length; index += 1) {
      value = values[index] || {};
      type = String(value.attachmentType || "file");
      if (!value.id || !/file/i.test(type)) {
        continue;
      }
      attachments.push({
        contentId: normalizeContentId(value.contentId),
        contentType: String(value.contentType || "").toLowerCase(),
        id: String(value.id),
        isInline: value.isInline !== false,
        name: String(value.name || "inline-image"),
        size: Number(value.size) || 0,
      });
    }
    return attachments;
  }

  function matchInlineAttachment(contentId, attachments) {
    var index;
    for (index = 0; index < attachments.length; index += 1) {
      if (
        attachments[index].isInline &&
        (attachments[index].contentId === contentId ||
          normalizeContentId(attachments[index].id) === contentId)
      ) {
        return attachments[index];
      }
    }
    return null;
  }

  function buildInlineImageFingerprintSource(contentIds, attachments) {
    var entries = [];
    var index;
    var attachment;
    for (index = 0; index < contentIds.length; index += 1) {
      attachment = matchInlineAttachment(contentIds[index], attachments);
      if (attachment) {
        entries.push(
          [
            contentIds[index],
            attachment.contentId || normalizeContentId(attachment.id),
            attachment.name,
            attachment.size,
            attachment.contentType,
          ].join("\u0000")
        );
      } else {
        entries.push(contentIds[index] + "\u0000missing");
      }
    }
    return entries.sort().join("\u0001");
  }

  function collectInlineImageContents(item, contentIds, attachments, callback, isCancelled) {
    var tasks = [];
    var refs = [];
    var uploads = [];
    var index;
    var attachment;
    var selectedMetadataBytes = 0;
    var decodedTotalBytes = 0;

    for (index = 0; index < contentIds.length; index += 1) {
      attachment = matchInlineAttachment(contentIds[index], attachments);
      if (
        attachment &&
        attachment.size <= MAX_INLINE_IMAGE_BYTES &&
        (!attachment.size ||
          selectedMetadataBytes + attachment.size <= MAX_INLINE_IMAGE_TOTAL_BYTES)
      ) {
        selectedMetadataBytes += attachment.size;
        tasks.push({ attachment: attachment, contentId: contentIds[index] });
      }
    }

    runWithConcurrency(
      tasks,
      4,
      function (task, next) {
        if (isCancelled()) {
          next();
          return;
        }
        if (!item.getAttachmentContentAsync) {
          next();
          return;
        }
        item.getAttachmentContentAsync(task.attachment.id, function (result) {
          var content;
          var format;
          if (!isSucceeded(result) || !result.value || !result.value.content) {
            next();
            return;
          }
          format = String(result.value.format || "").toLowerCase();
          if (format && format.indexOf("base64") < 0) {
            next();
            return;
          }
          content = String(result.value.content).replace(/\s/g, "");
          var decodedBytes = getDecodedBase64Length(content);
          if (
            decodedBytes > MAX_INLINE_IMAGE_BYTES ||
            decodedTotalBytes + decodedBytes > MAX_INLINE_IMAGE_TOTAL_BYTES
          ) {
            next();
            return;
          }
          decodedTotalBytes += decodedBytes;
          sha256Base64(content, function (hashError, hash) {
            if (!hashError && hash) {
              refs.push({ contentId: task.contentId, sha256: hash });
              uploads.push({
                contentBase64: content,
                contentType:
                  task.attachment.contentType || inferInlineImageContentType(task.attachment.name),
                name: task.attachment.name,
                sha256: hash,
              });
            }
            next();
          });
        });
      },
      function () {
        callback(refs, uploads);
      }
    );
  }

  function runWithConcurrency(tasks, limit, worker, callback) {
    var nextIndex = 0;
    var active = 0;
    var finished = 0;

    if (!tasks.length) {
      callback();
      return;
    }

    function schedule() {
      var task;
      while (active < limit && nextIndex < tasks.length) {
        task = tasks[nextIndex];
        nextIndex += 1;
        active += 1;
        worker(task, function () {
          active -= 1;
          finished += 1;
          if (finished === tasks.length) {
            callback();
            return;
          }
          schedule();
        });
      }
    }
    schedule();
  }

  function sha256Base64(value, callback) {
    var binary;
    var bytes;
    var index;
    try {
      binary = atob(value);
      bytes = new Uint8Array(binary.length);
      for (index = 0; index < binary.length; index += 1) {
        bytes[index] = binary.charCodeAt(index);
      }
    } catch (error) {
      callback(error, "");
      return;
    }
    sha256Bytes(bytes, callback);
  }

  function getDecodedBase64Length(value) {
    var normalized = String(value || "").replace(/\s/g, "");
    var padding = 0;
    if (/==$/.test(normalized)) {
      padding = 2;
    } else if (/=$/.test(normalized)) {
      padding = 1;
    }
    return Math.max(0, Math.floor((normalized.length * 3) / 4) - padding);
  }

  function sha256Bytes(bytes, callback) {
    var cryptoProvider = window.crypto;
    var subtle;
    var operation;
    var digestInput = bytes;
    if (bytes && bytes.buffer) {
      digestInput = bytes.buffer;
    }
    if (!cryptoProvider || !cryptoProvider.subtle) {
      cryptoProvider = window.msCrypto;
    }
    subtle = cryptoProvider && cryptoProvider.subtle;
    if (!subtle || !subtle.digest) {
      callback(new Error("SHA-256 is unavailable."), "");
      return;
    }
    try {
      operation = subtle.digest("SHA-256", digestInput);
    } catch (error) {
      void error;
      try {
        operation = subtle.digest({ name: "SHA-256" }, digestInput);
      } catch (legacyError) {
        callback(legacyError, "");
        return;
      }
    }
    if (operation && typeof operation.then === "function") {
      operation.then(
        function (digest) {
          finishSha256(digest, callback);
        },
        function (error) {
          callback(error, "");
        }
      );
      return;
    }
    if (!operation || typeof operation !== "object") {
      callback(new Error("SHA-256 failed."), "");
      return;
    }
    operation.oncomplete = function (event) {
      finishSha256(event && event.target && event.target.result, callback);
    };
    operation.onerror = function () {
      callback(new Error("SHA-256 failed."), "");
    };
  }

  function finishSha256(digest, callback) {
    var values;
    var hex = "";
    var index;
    if (!digest) {
      callback(new Error("SHA-256 returned no result."), "");
      return;
    }
    values = new Uint8Array(digest);
    for (index = 0; index < values.length; index += 1) {
      hex += ("0" + values[index].toString(16)).slice(-2);
    }
    callback(null, hex);
  }

  function utf8Bytes(value) {
    var encoded = unescape(encodeURIComponent(String(value || "")));
    var bytes = new Uint8Array(encoded.length);
    var index;
    for (index = 0; index < encoded.length; index += 1) {
      bytes[index] = encoded.charCodeAt(index);
    }
    return bytes;
  }

  function inferInlineImageContentType(filename) {
    var extension = String(filename || "")
      .split(".")
      .pop()
      .toLowerCase();
    if (extension === "png") {
      return "image/png";
    }
    if (extension === "jpg" || extension === "jpeg") {
      return "image/jpeg";
    }
    if (extension === "gif") {
      return "image/gif";
    }
    if (extension === "webp") {
      return "image/webp";
    }
    return "application/octet-stream";
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

  function cancelEmailAttachmentPrefetch(prefetchKey, backgroundSessionId) {
    if (!prefetchKey) {
      return;
    }
    sendJsonRequestMethod(
      "DELETE",
      EMAIL_ATTACHMENT_PREFETCH_BASE_URL + encodeURIComponent(prefetchKey),
      { backgroundSessionId: backgroundSessionId || "" },
      SEND_REQUEST_TIMEOUT_MS,
      function () {}
    );
  }

  function sendJsonRequest(url, payload, timeoutMs, callback) {
    sendJsonRequestMethod("POST", url, payload, timeoutMs, callback);
  }

  function sendJsonRequestMethod(method, url, payload, timeoutMs, callback) {
    var xhr = new XMLHttpRequest();
    var completed = false;
    var timeout = setTimeout(
      function () {
        finish(new Error("Halo attachment request timed out."));
      },
      Math.max(1, timeoutMs || SEND_REQUEST_TIMEOUT_MS)
    );

    function finish(error, result) {
      if (completed) {
        return;
      }
      completed = true;
      clearTimeout(timeout);
      callback(error, result);
    }

    xhr.open(method, url, true);
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
        parsed = { ok: xhr.status >= 200 && xhr.status < 300, status: "failed" };
      }
      finish(null, parsed);
    };
    xhr.onerror = function () {
      finish(new Error("Halo attachment request failed."));
    };
    xhr.send(JSON.stringify(payload || {}));
  }

  function sendRuntimeDiagnostic(stage, startedAt, details) {
    var xhr;
    var payload = {
      elapsedMs: Math.max(0, Date.now() - startedAt),
      stage: stage,
    };
    var allowedKeys = [
      "attachmentCount",
      "attachmentError",
      "attemptCount",
      "failedCount",
      "hasBackgroundSession",
      "hasMailboxIdentity",
      "includeAttachments",
      "inlineImageCount",
      "outcome",
      "skippedCount",
      "statusCode",
      "uploadedCount",
    ];
    var index;
    var key;

    if (details) {
      for (index = 0; index < allowedKeys.length; index += 1) {
        key = allowedKeys[index];
        if (details[key] !== undefined) {
          payload[key] = details[key];
        }
      }
    }

    try {
      xhr = new XMLHttpRequest();
      xhr.open("POST", SEND_DIAGNOSTIC_URL, true);
      xhr.withCredentials = true;
      xhr.timeout = 1000;
      xhr.setRequestHeader("Accept", "application/json");
      xhr.setRequestHeader("Content-Type", "application/json");
      xhr.onreadystatechange = function () {};
      xhr.onerror = function () {};
      xhr.ontimeout = function () {};
      xhr.send(JSON.stringify(payload));
    } catch (error) {
      void error;
    }
  }

  function getDraftSaveOutcome(saveError) {
    if (saveError) {
      return "failed";
    }
    return "mismatch";
  }

  function getSendAssetDiagnostics(email) {
    var attachmentCount = 0;
    var includeAttachments = Boolean(email && email.includeEmailAttachments);
    var inlineImageCount = 0;
    if (email && email.emailAttachmentSummary) {
      attachmentCount = Number(email.emailAttachmentSummary.selected) || 0;
    }
    if (email && email.inlineImageRefs && email.inlineImageRefs.length) {
      inlineImageCount = email.inlineImageRefs.length;
    }
    return {
      attachmentCount: attachmentCount,
      includeAttachments: includeAttachments,
      inlineImageCount: inlineImageCount,
      outcome: "ok",
    };
  }

  function getSendRequestOutcome(error, result) {
    if (error) {
      if (
        String(error.message || "")
          .toLowerCase()
          .indexOf("timed out") >= 0
      ) {
        return "timeout";
      }
      return "network-error";
    }
    if (!result) {
      return "missing-response";
    }
    if (result.status) {
      return String(result.status);
    }
    if (result.ok) {
      return "ok";
    }
    return "failed";
  }

  function applyAutomaticRecovery(email, recovery) {
    email.mappingRecoveryValidated = true;
    email.recoveredComposeAttachId = recovery.recoveredComposeAttachId || "";
    email.ticketId = recovery.ticketId || "";
    email.ticketNumber = recovery.ticketNumber || "";
    email.actionMode = normalizeActionMode(recovery.actionMode);
  }

  function recoverAutomaticMapping(item, email, callback) {
    var recoveryStartedAt = Date.now();
    var flightKey = String(email.itemId || "") + "|" + recoveryIdentifierFingerprint(email);
    var existing = recoveryFlights[flightKey];
    if (existing) {
      existing.push(callback);
      return;
    }
    recoveryFlights[flightKey] = [callback];
    sendRuntimeDiagnostic("recovery-start", recoveryStartedAt, { outcome: "started" });
    var finished = false;
    var timer = setTimeout(function () {
      finish({ status: "no-match" });
    }, RECOVERY_TIMEOUT_MS);

    function finish(result) {
      var callbacks;
      var index;
      if (finished) {
        return;
      }
      finished = true;
      clearTimeout(timer);
      callbacks = recoveryFlights[flightKey] || [];
      delete recoveryFlights[flightKey];
      if (result && result.status === "matched") {
        saveRecoveryCache(item, email, result);
      }
      sendRuntimeDiagnostic("recovery-complete", recoveryStartedAt, {
        outcome: result && result.status === "matched" ? "matched" : "no-match",
      });
      for (index = 0; index < callbacks.length; index += 1) {
        callbacks[index](result);
      }
    }

    readRecoveryCache(item, email, function (cached) {
      collectExactChainIdentifiers(item, email, function (identifiers) {
        var remaining = 2;
        var cachedId = cached && cached.recoveredComposeAttachId;
        function branchComplete(candidates) {
          var values = candidates || [];
          if (cachedId && values.indexOf(cachedId) < 0) {
            values.unshift(cachedId);
          }
          sendRuntimeDiagnostic("recovery-branch", recoveryStartedAt, {
            candidateCount: values.length,
            outcome: values.length ? "ok" : "no-match",
          });
          validateRecoveryCandidates(email, values, function (result) {
            if (finished) {
              return;
            }
            if (result && result.status === "matched") {
              finish(result);
              return;
            }
            remaining -= 1;
            if (!remaining) {
              finish({ status: "no-match" });
            }
          });
        }
        resolveComposeIdsFromHeaders(item, identifiers, branchComplete);
        resolveComposeIdsFromSentProperties(identifiers, branchComplete);
      });
    });
  }

  function validateRecoveryCandidates(email, candidates, callback) {
    var payload = {
      backgroundSessionId: email.backgroundSessionId,
      composeAttachIds: uniqueOpaqueComposeIds(candidates),
      conversationId: email.conversationId,
      inReplyToMessageIds: email.inReplyToMessageIds,
      internetMessageId: email.internetMessageId,
      itemId: email.itemId,
      mailboxEmail: email.mailboxEmail,
      referenceMessageIds: email.referenceMessageIds,
    };
    sendJsonRequest(RECOVER_MAPPING_URL, payload, RECOVERY_TIMEOUT_MS, function (error, result) {
      var candidateIndex;
      if (error || !result || result.status !== "matched") {
        callback({ status: "no-match" });
        return;
      }
      candidateIndex = Number(result.candidateIndex);
      result.recoveredComposeAttachId =
        candidateIndex >= 0 && candidateIndex < payload.composeAttachIds.length
          ? payload.composeAttachIds[candidateIndex]
          : "";
      callback(result);
    });
  }

  function collectExactChainIdentifiers(item, email, callback) {
    var identifiers = uniqueMessageIds(
      (email.inReplyToMessageIds || []).concat(email.referenceMessageIds || [])
    );
    // item.inReplyTo identifies only the immediate parent. For replies to an
    // incoming response, the originally sent mapped message is normally found
    // only in PR_INTERNET_REFERENCES. Always enrich the chain from the saved
    // draft when EWS is available instead of treating inReplyTo as complete.
    if (!email.itemId || !canMakeEwsRequest()) {
      callback(identifiers);
      return;
    }
    makeEwsRequest(buildDraftChainGetItemRequest(email.itemId), function (error, response) {
      if (!error) {
        identifiers = uniqueMessageIds(
          identifiers.concat(extractMessageIdsFromExtendedValues(response))
        );
        email.inReplyToMessageIds = uniqueMessageIds(
          (email.inReplyToMessageIds || []).concat(identifiers.slice(0, 1))
        );
        email.referenceMessageIds = uniqueMessageIds(
          (email.referenceMessageIds || []).concat(identifiers)
        );
      }
      callback(identifiers);
    });
  }

  function resolveComposeIdsFromHeaders(item, identifiers, callback) {
    var direct = [];
    var waiting = 1;
    function finishPart(values) {
      direct = uniqueOpaqueComposeIds(direct.concat(values || []));
      waiting -= 1;
      if (!waiting) {
        callback(direct);
      }
    }
    if (item && item.internetHeaders && item.internetHeaders.getAsync) {
      waiting += 1;
      item.internetHeaders.getAsync([COMPOSE_ATTACH_HEADER_NAME], function (result) {
        var value = isSucceeded(result) && result.value && result.value[COMPOSE_ATTACH_HEADER_NAME];
        finishPart(value ? [String(value)] : []);
      });
    }
    if (!identifiers.length || !canMakeEwsRequest()) {
      finishPart([]);
      return;
    }
    findEwsItemsByInternetMessageIds(identifiers, ["inbox", "sentitems"], function (itemIds) {
      if (!itemIds.length) {
        finishPart([]);
        return;
      }
      makeEwsRequest(buildHeaderGetItemRequest(itemIds), function (error, response) {
        finishPart(error ? [] : extractComposeIdsFromHeaders(response));
      });
    });
  }

  function resolveComposeIdsFromSentProperties(identifiers, callback) {
    if (!identifiers.length || !canMakeEwsRequest()) {
      callback([]);
      return;
    }
    findEwsItemsByInternetMessageIds(identifiers, ["sentitems"], function (itemIds) {
      if (!itemIds.length) {
        callback([]);
        return;
      }
      makeEwsRequest(buildCustomPropertyGetItemRequest(itemIds), function (error, response) {
        callback(error ? [] : extractComposeIdsFromCustomProperties(response));
      });
    });
  }

  function findEwsItemsByInternetMessageIds(identifiers, folders, callback) {
    makeEwsRequest(
      buildFindItemRequest(identifiers.slice(0, 12), folders),
      function (error, response) {
        callback(error ? [] : extractEwsItemIds(response).slice(0, 12));
      }
    );
  }

  function canMakeEwsRequest() {
    return Boolean(Office.context.mailbox && Office.context.mailbox.makeEwsRequestAsync);
  }

  function makeEwsRequest(request, callback) {
    if (!canMakeEwsRequest()) {
      callback(new Error("EWS is unavailable."), "");
      return;
    }
    Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
      if (!isSucceeded(result) || !result.value || String(result.value).length > 1024 * 1024) {
        callback(new Error("EWS metadata lookup failed."), "");
        return;
      }
      callback(null, String(result.value));
    });
  }

  function buildEwsEnvelope(body) {
    return (
      '<?xml version="1.0" encoding="utf-8"?>' +
      '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
      'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
      'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">' +
      '<soap:Header><t:RequestServerVersion Version="Exchange2013"/></soap:Header>' +
      "<soap:Body>" +
      body +
      "</soap:Body></soap:Envelope>"
    );
  }

  function buildDraftChainGetItemRequest(itemId) {
    return buildEwsEnvelope(
      "<m:GetItem><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape>" +
        "<t:AdditionalProperties>" +
        '<t:ExtendedFieldURI PropertyTag="0x1042" PropertyType="String"/>' +
        '<t:ExtendedFieldURI PropertyTag="0x1039" PropertyType="String"/>' +
        '</t:AdditionalProperties></m:ItemShape><m:ItemIds><t:ItemId Id="' +
        escapeXml(itemId) +
        '"/></m:ItemIds></m:GetItem>'
    );
  }

  function buildFindItemRequest(identifiers, folders) {
    var restrictions = identifiers.map(function (id) {
      return (
        '<t:IsEqualTo><t:FieldURI FieldURI="message:InternetMessageId"/>' +
        '<t:FieldURIOrConstant><t:Constant Value="' +
        escapeXml(id) +
        '"/></t:FieldURIOrConstant></t:IsEqualTo>'
      );
    });
    var restriction =
      restrictions.length === 1 ? restrictions[0] : "<t:Or>" + restrictions.join("") + "</t:Or>";
    var parentFolders = folders
      .map(function (folder) {
        return '<t:DistinguishedFolderId Id="' + folder + '"/>';
      })
      .join("");
    return buildEwsEnvelope(
      '<m:FindItem Traversal="Shallow"><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape></m:ItemShape>' +
        "<m:Restriction>" +
        restriction +
        "</m:Restriction><m:ParentFolderIds>" +
        parentFolders +
        "</m:ParentFolderIds></m:FindItem>"
    );
  }

  function buildHeaderGetItemRequest(itemIds) {
    return buildMetadataGetItemRequest(
      itemIds,
      '<t:FieldURI FieldURI="item:InternetMessageHeaders"/>'
    );
  }

  function buildCustomPropertyGetItemRequest(itemIds) {
    return buildMetadataGetItemRequest(
      itemIds,
      '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="' +
        COMPOSE_CUSTOM_PROPERTY_NAME +
        '" PropertyType="String"/>'
    );
  }

  function buildMetadataGetItemRequest(itemIds, property) {
    return buildEwsEnvelope(
      "<m:GetItem><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties>" +
        property +
        "</t:AdditionalProperties></m:ItemShape><m:ItemIds>" +
        itemIds
          .map(function (id) {
            return '<t:ItemId Id="' + escapeXml(id) + '"/>';
          })
          .join("") +
        "</m:ItemIds></m:GetItem>"
    );
  }

  function extractEwsItemIds(xml) {
    var values = [];
    var regex = /<(?:\w+:)?ItemId\b[^>]*\bId="([^"]+)"/gi;
    var match;
    while ((match = regex.exec(xml)) && values.length < 12) {
      values.push(unescapeXml(match[1]));
    }
    return values;
  }

  function extractMessageIds(value) {
    var values = [];
    var match;
    var regex = /<[^<>\s]+>/g;
    value = unescapeXml(String(value || ""));
    while ((match = regex.exec(value)) && values.length < 24) {
      values.push(match[0]);
    }
    return uniqueMessageIds(values);
  }

  function extractMessageIdsFromExtendedValues(xml) {
    var values = [];
    var regex = /<(?:\w+:)?Value>([\s\S]*?)<\/(?:\w+:)?Value>/gi;
    var match;
    while ((match = regex.exec(String(xml || ""))) && values.length < 24) {
      values = values.concat(extractMessageIds(match[1]));
    }
    return uniqueMessageIds(values);
  }

  function extractComposeIdsFromHeaders(xml) {
    var values = [];
    var regex =
      /<(?:\w+:)?InternetMessageHeader\b[^>]*HeaderName="X-Halo-Compose-Id"[^>]*>([\s\S]*?)<\/(?:\w+:)?InternetMessageHeader>/gi;
    var match;
    while ((match = regex.exec(xml)) && values.length < 12) {
      values.push(unescapeXml(match[1]).trim());
    }
    return uniqueOpaqueComposeIds(values);
  }

  function extractComposeIdsFromCustomProperties(xml) {
    var values = [];
    var regex =
      /<(?:\w+:)?ExtendedProperty>[\s\S]*?PropertyName="cecp-55bbcff2-8191-4411-aec6-f9d2f9b4b5e8"[\s\S]*?<(?:\w+:)?Value>([\s\S]*?)<\/(?:\w+:)?Value>[\s\S]*?<\/(?:\w+:)?ExtendedProperty>/gi;
    var match;
    while ((match = regex.exec(xml)) && values.length < 12) {
      try {
        var properties = JSON.parse(unescapeXml(match[1]));
        var markerValue = properties && properties[COMPOSE_ATTACH_STORAGE_KEY];
        var marker = typeof markerValue === "string" ? JSON.parse(markerValue) : markerValue;
        if (marker && marker.composeAttachId) {
          values.push(String(marker.composeAttachId));
        }
      } catch (error) {
        void error;
      }
    }
    return uniqueOpaqueComposeIds(values);
  }

  function uniqueOpaqueComposeIds(values) {
    var result = [];
    var seen = {};
    var index;
    for (index = 0; index < (values || []).length && result.length < 12; index += 1) {
      var value = String(values[index] || "");
      if (value && value.length <= 200 && /^[A-Za-z0-9_-]+$/.test(value) && !seen[value]) {
        seen[value] = true;
        result.push(value);
      }
    }
    return result;
  }

  function uniqueMessageIds(values) {
    var result = [];
    var seen = {};
    var index;
    for (index = 0; index < (values || []).length && result.length < 24; index += 1) {
      var value = String(values[index] || "").trim();
      if (value && value.charAt(0) !== "<") {
        value = "<" + value.replace(/^<|>$/g, "") + ">";
      }
      var key = value.toLowerCase();
      if (value && !seen[key]) {
        seen[key] = true;
        result.push(value);
      }
    }
    return result;
  }

  function escapeXml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function unescapeXml(value) {
    return String(value || "")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, "&");
  }

  function recoveryIdentifierFingerprint(email) {
    var values = uniqueMessageIds(
      (email.inReplyToMessageIds || []).concat(email.referenceMessageIds || [])
    );
    return String(email.conversationId || "") + "|" + values.sort().join("|");
  }

  function readRecoveryCache(item, email, callback) {
    if (!item || !item.sessionData || !item.sessionData.getAsync) {
      callback(null);
      return;
    }
    item.sessionData.getAsync(MAPPING_RECOVERY_STORAGE_KEY, function (result) {
      var value = null;
      try {
        value = isSucceeded(result) ? JSON.parse(result.value || "null") : null;
      } catch (error) {
        void error;
      }
      var currentFingerprint = recoveryIdentifierFingerprint(email);
      var hasCurrentChainIds = Boolean(
        (email.inReplyToMessageIds && email.inReplyToMessageIds.length) ||
        (email.referenceMessageIds && email.referenceMessageIds.length)
      );
      if (
        !value ||
        value.version !== 1 ||
        value.draftItemId !== email.itemId ||
        (hasCurrentChainIds && value.chainFingerprint !== currentFingerprint)
      ) {
        callback(null);
        return;
      }
      callback(value);
    });
  }

  function saveRecoveryCache(item, email, result) {
    if (!item || !item.sessionData || !item.sessionData.setAsync) {
      return;
    }
    item.sessionData.setAsync(
      MAPPING_RECOVERY_STORAGE_KEY,
      JSON.stringify({
        version: 1,
        draftItemId: email.itemId,
        chainFingerprint: recoveryIdentifierFingerprint(email),
        recoveredComposeAttachId: result.recoveredComposeAttachId || "",
      }),
      function () {}
    );
  }

  function sendRequest(url, email, callback) {
    var xhr = new XMLHttpRequest();
    var completed = false;
    var timeout = setTimeout(function () {
      finish(new Error("Halo send auto-attach timed out."));
    }, SEND_REQUEST_TIMEOUT_MS);

    function finish(error, result) {
      if (completed) {
        return;
      }

      completed = true;
      clearTimeout(timeout);
      callback(error, result, Number(xhr.status) || 0);
    }

    xhr.open("POST", url, true);
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

  function createSendEventCompletion(event, onComplete) {
    var completed = false;

    return function (options) {
      if (completed) {
        return;
      }

      completed = true;
      onComplete();
      event.completed(options);
    };
  }

  function completeAllow(complete) {
    complete({ allowEvent: true });
  }

  function completeWithHaloWarning(complete, result, failure) {
    var creatingTicket = isCreateTicketMarker(result);
    var advice = "Send anyway or try again.";
    var errorPrefix;
    var failureMessage = "";
    var ticketLabel = result.ticketNumber || result.ticketId;
    if (failure && failure.message) {
      failureMessage = String(failure.message).toLowerCase();
    }
    if (result && result.status === "no-session") {
      advice = "Reopen the Halo add-in to renew its send session, then try again.";
    } else if (result && result.status === "attachments-not-ready") {
      advice = "Open the Halo pane, finish attachment preparation, and retry Send.";
    } else if (failure === "draft") {
      advice =
        "The saved draft changed. Reopen the Halo add-in, select the ticket again, and retry.";
    } else if (failure === "compose") {
      advice = "Outlook could not read the final message. Review the draft and try again.";
    } else if (failure === "timeout" || failureMessage.indexOf("timed out") >= 0) {
      advice = "The Halo request timed out. Try again, or send anyway.";
    } else if (failureMessage.indexOf("request failed") >= 0) {
      advice = "The Halo service could not be reached. Try again, or send anyway.";
    }
    if (!ticketLabel) {
      ticketLabel = "the mapped Halo ticket";
      if (creatingTicket) {
        ticketLabel = result.ticketTypeName || "a Halo ticket";
      }
    }
    errorPrefix = "Could not add this email to Halo ticket " + ticketLabel + ". ";
    if (creatingTicket) {
      errorPrefix = "Could not create " + ticketLabel + " from this email. ";
    }
    complete({
      allowEvent: false,
      errorMessage: errorPrefix + advice,
    });
  }

  Office.actions.associate("onHaloMessageSend", onHaloMessageSend);
  Office.actions.associate("onHaloMessageAttachmentsChanged", onHaloMessageAttachmentsChanged);
  sendRuntimeDiagnostic("runtime-loaded", Date.now(), { outcome: "ok" });
})();
