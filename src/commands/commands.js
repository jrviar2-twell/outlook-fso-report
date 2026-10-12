/**
 * Report Email to FSO – function command for New Outlook / Outlook on the web / classic Outlook.
 *
 * When the user clicks "Report Email to FSO" on a message:
 *  1. Opens a new compose form addressed to jrviar@tracewell.com
 *  2. Pre-fills subject and body with original message info
 *  3. Attaches the original message as an Outlook item
 *  4. Lets the user review / send (or they can just click Send)
 */

/* global Office */

Office.onReady(() => {
  // Required. Function commands still need Office.onReady.
});

/**
 * Main entry point called from the ribbon button.
 * @param {Office.AddinCommands.Event} event
 */
function reportEmailToFSO(event) {
  try {
    const item = Office.context.mailbox.item;

    if (!item) {
      showNotification("error", "No email selected. Please select a message first.");
      event.completed();
      return;
    }

    // Gather details from the current message
    const subject = item.subject || "(no subject)";
    let fromDisplay = "unknown";

    if (item.from) {
      const name = item.from.displayName || "";
      const addr = item.from.emailAddress || "";
      fromDisplay = name ? `${name} <${addr}>` : addr;
    }

    const reportedAt = new Date().toLocaleString();

    const htmlBody =
      "This email was reported by the user.<br/><br/>" +
      "<b>Original Subject:</b> " + escapeHtml(subject) + "<br/>" +
      "<b>Original From:</b> " + escapeHtml(fromDisplay) + "<br/>" +
      "<b>Reported at:</b> " + escapeHtml(reportedAt);

    // Open a new message form with the original item attached
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: ["jrviar@tracewell.com"],
      subject: "Reported Email: " + subject,
      htmlBody: htmlBody,
      attachments: [
        {
          type: "item",
          name: "Reported Email",
          itemId: item.itemId
        }
      ]
    });

    // Optional: show a brief success note in the original message
    showNotification("informationalMessage", "Report form opened. Review and click Send.");
  } catch (err) {
    console.error("reportEmailToFSO error:", err);
    showNotification("error", "Error reporting email: " + (err.message || err));
  } finally {
    // Always signal that the function command has finished
    event.completed();
  }
}

/**
 * Helper – show a notification banner on the current item.
 */
function showNotification(type, message) {
  try {
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "fsoReportStatus",
      {
        type: type,          // "informationalMessage" | "errorMessage" | "progressIndicator"
        message: message,
        icon: "Icon.16",
        persistent: false
      }
    );
  } catch (e) {
    // Notification API not available in every host – ignore
    console.warn("Could not show notification:", e);
  }
}

/**
 * Simple HTML escape for values we put into the report body.
 */
function escapeHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Register the function so the manifest can call it
Office.actions.associate("reportEmailToFSO", reportEmailToFSO);
