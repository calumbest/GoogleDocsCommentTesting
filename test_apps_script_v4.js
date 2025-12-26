/**
 * Test 3d: Pass BOTH anchor AND quotedFileContent
 *
 * Theory: The full anchored behavior requires both:
 * - anchor: "kix.xxxxx"
 * - quotedFileContent: {"mimeType":"text/html","value":"the highlighted text"}
 */

/**
 * Test 3d: Create comment with anchor + quotedFileContent
 */
function test_createWithFullAnchor() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  // Use the anchor from the original manual comment
  const anchor = "kix.hkaqsaj0l7p3";
  const quotedFileContent = {
    mimeType: "text/html",
    value: "ipsum"
  };

  Logger.log("Creating comment with:");
  Logger.log("  anchor: " + anchor);
  Logger.log("  quotedFileContent: " + JSON.stringify(quotedFileContent));

  const comment = {
    content: "TEST 3d: Using BOTH anchor AND quotedFileContent",
    anchor: anchor,
    quotedFileContent: quotedFileContent
  };

  try {
    const result = Drive.Comments.create(comment, fileId, {
      fields: "id,content,anchor,quotedFileContent"
    });

    Logger.log("\nResult:");
    Logger.log("  ID: " + result.id);
    Logger.log("  Anchor: " + result.anchor);
    Logger.log("  QuotedFileContent: " + JSON.stringify(result.quotedFileContent));

    DocumentApp.getUi().alert(
      "Comment created!\n\n" +
      "Anchor: " + (result.anchor || "NONE") + "\n" +
      "QuotedFileContent: " + JSON.stringify(result.quotedFileContent) + "\n\n" +
      "Check the document:\n" +
      "1. Does comment panel show 'Tab 1 · ipsum'?\n" +
      "2. Is there yellow highlighting on 'ipsum' in the doc?"
    );
  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}

/**
 * Find the original manual comment and clone it exactly
 */
function test_cloneManualComment() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  Logger.log("Looking for original manual comment with quotedFileContent...");

  try {
    const comments = Drive.Comments.list(fileId, {
      fields: "comments(id,content,anchor,quotedFileContent)"
    });

    // Find a comment that has BOTH anchor and quotedFileContent
    let sourceComment = null;
    for (let i = 0; i < comments.comments.length; i++) {
      const c = comments.comments[i];
      if (c.anchor && c.quotedFileContent && c.quotedFileContent.value) {
        sourceComment = c;
        Logger.log("Found source comment:");
        Logger.log("  Content: " + c.content);
        Logger.log("  Anchor: " + c.anchor);
        Logger.log("  QuotedFileContent: " + JSON.stringify(c.quotedFileContent));
        break;
      }
    }

    if (!sourceComment) {
      DocumentApp.getUi().alert(
        "No comment with both anchor AND quotedFileContent found.\n\n" +
        "Your original manual comment may have been deleted.\n" +
        "Please create a new comment manually, then try again."
      );
      return;
    }

    // Clone it
    Logger.log("\nCloning comment...");
    const newComment = {
      content: "TEST 3d-clone: Exact clone of manual comment",
      anchor: sourceComment.anchor,
      quotedFileContent: sourceComment.quotedFileContent
    };

    const result = Drive.Comments.create(newComment, fileId, {
      fields: "id,content,anchor,quotedFileContent"
    });

    Logger.log("\nClone result:");
    Logger.log("  ID: " + result.id);
    Logger.log("  Anchor: " + result.anchor);
    Logger.log("  QuotedFileContent: " + JSON.stringify(result.quotedFileContent));

    DocumentApp.getUi().alert(
      "Cloned comment created!\n\n" +
      "Source anchor: " + sourceComment.anchor + "\n" +
      "Result anchor: " + (result.anchor || "NONE") + "\n\n" +
      "Check if the clone behaves like the original."
    );
  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}

/**
 * List comments showing which have full anchoring
 */
function test_showAnchorStatus() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  try {
    const comments = Drive.Comments.list(fileId, {
      fields: "comments(id,content,anchor,quotedFileContent)"
    });

    Logger.log("Comment Anchor Status:\n");

    comments.comments.forEach(function(c, i) {
      const hasAnchor = c.anchor ? "YES" : "NO";
      const hasQuoted = (c.quotedFileContent && c.quotedFileContent.value) ? "YES" : "NO";
      const status = (hasAnchor === "YES" && hasQuoted === "YES") ? "FULL" :
                     (hasAnchor === "YES") ? "PARTIAL" : "NONE";

      Logger.log((i+1) + ". " + status + " anchoring");
      Logger.log("   Content: " + c.content.substring(0, 40) + "...");
      Logger.log("   Anchor: " + (c.anchor || "none"));
      Logger.log("   Quoted: " + (c.quotedFileContent ? c.quotedFileContent.value : "none"));
      Logger.log("");
    });

    DocumentApp.getUi().alert("Check View > Logs for anchor status of all comments.");
  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}
