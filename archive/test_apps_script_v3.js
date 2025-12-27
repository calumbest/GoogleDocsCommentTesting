/**
 * Test 3b & 3c: Anchor Experiments
 *
 * SETUP: Same as v2 - need Drive API service enabled
 * 1. Extensions > Apps Script
 * 2. Services > Add > Drive API
 * 3. Paste this code, save, run tests
 */

/**
 * Test 3b: Try to create comment with a kix anchor string
 * Uses the anchor format we captured from a real comment
 */
function test_createWithKixAnchor() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  // Try passing a kix-style anchor (captured from real comment)
  // This is the format Google uses internally
  const kixAnchor = "kix.hkaqsaj0l7p3";

  Logger.log("Attempting to create comment with kix anchor: " + kixAnchor);

  const comment = {
    content: "TEST 3b: Trying to use kix anchor format directly",
    anchor: kixAnchor
  };

  try {
    const result = Drive.Comments.create(comment, fileId, {
      fields: "id,content,anchor,quotedFileContent"
    });

    Logger.log("Result:");
    Logger.log("  ID: " + result.id);
    Logger.log("  Anchor in response: " + result.anchor);
    Logger.log("  QuotedFileContent: " + JSON.stringify(result.quotedFileContent));

    const anchored = result.anchor ? "YES - " + result.anchor : "NO - undefined";
    DocumentApp.getUi().alert(
      "Comment created!\n\n" +
      "Anchored? " + anchored + "\n\n" +
      "Check the document to see if it's attached to 'ipsum'"
    );
  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}

/**
 * Test 3c: Get an existing anchor, then create new comment with same anchor
 * Tests if we can "clone" a comment's position
 */
function test_reuseExistingAnchor() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  // First, get existing comments
  Logger.log("Step 1: Fetching existing comments...");

  let existingAnchor = null;
  let existingQuoted = null;

  try {
    const comments = Drive.Comments.list(fileId, {
      fields: "comments(id,content,anchor,quotedFileContent)"
    });

    // Find a comment that HAS an anchor (manually created one)
    for (let i = 0; i < comments.comments.length; i++) {
      const c = comments.comments[i];
      if (c.anchor && c.anchor.startsWith("kix.")) {
        existingAnchor = c.anchor;
        existingQuoted = c.quotedFileContent;
        Logger.log("Found anchored comment:");
        Logger.log("  Content: " + c.content);
        Logger.log("  Anchor: " + c.anchor);
        Logger.log("  QuotedFileContent: " + JSON.stringify(c.quotedFileContent));
        break;
      }
    }

    if (!existingAnchor) {
      DocumentApp.getUi().alert(
        "No anchored comments found!\n\n" +
        "Please create a comment manually in the Google Docs UI first, " +
        "then run this test again."
      );
      return;
    }

  } catch (e) {
    Logger.log("Error fetching comments: " + e.toString());
    DocumentApp.getUi().alert("Error fetching comments: " + e.toString());
    return;
  }

  // Now try to create a NEW comment with the SAME anchor
  Logger.log("\nStep 2: Creating new comment with stolen anchor...");
  Logger.log("  Using anchor: " + existingAnchor);

  const newComment = {
    content: "TEST 3c: This comment is trying to REUSE anchor " + existingAnchor,
    anchor: existingAnchor
  };

  // Also try with quotedFileContent
  if (existingQuoted) {
    newComment.quotedFileContent = existingQuoted;
  }

  try {
    const result = Drive.Comments.create(newComment, fileId, {
      fields: "id,content,anchor,quotedFileContent"
    });

    Logger.log("Result:");
    Logger.log("  ID: " + result.id);
    Logger.log("  Anchor in response: " + result.anchor);
    Logger.log("  QuotedFileContent: " + JSON.stringify(result.quotedFileContent));

    const anchored = result.anchor ? "YES - " + result.anchor : "NO - undefined";
    DocumentApp.getUi().alert(
      "New comment created with reused anchor!\n\n" +
      "Original anchor: " + existingAnchor + "\n" +
      "Result anchor: " + anchored + "\n\n" +
      "Check the document - is the new comment anchored to the same text?"
    );
  } catch (e) {
    Logger.log("ERROR creating comment: " + e.toString());
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}

/**
 * Utility: List all comments with full details
 */
function test_listAllComments() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  try {
    const comments = Drive.Comments.list(fileId, {
      fields: "comments(id,content,author,anchor,quotedFileContent,resolved)"
    });

    if (!comments.comments || comments.comments.length === 0) {
      Logger.log("No comments found");
      DocumentApp.getUi().alert("No comments found in this document.");
      return;
    }

    Logger.log("Found " + comments.comments.length + " comments:\n");

    comments.comments.forEach(function(c, i) {
      Logger.log("=== Comment " + (i+1) + " ===");
      Logger.log("ID: " + c.id);
      Logger.log("Content: " + c.content);
      Logger.log("Author: " + (c.author ? c.author.displayName : "unknown"));
      Logger.log("Anchor: " + (c.anchor || "NONE"));
      Logger.log("QuotedFileContent: " + JSON.stringify(c.quotedFileContent));
      Logger.log("Resolved: " + c.resolved);
      Logger.log("");
    });

    DocumentApp.getUi().alert(
      "Found " + comments.comments.length + " comments.\n\n" +
      "Check View > Logs for full details."
    );
  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}

/**
 * Cleanup: Delete all unanchored comments (API-created ones)
 */
function test_deleteUnanchoredComments() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  try {
    const comments = Drive.Comments.list(fileId, {
      fields: "comments(id,anchor)"
    });

    let deleted = 0;
    comments.comments.forEach(function(c) {
      if (!c.anchor) {
        Drive.Comments.remove(fileId, c.id);
        Logger.log("Deleted unanchored comment: " + c.id);
        deleted++;
      }
    });

    DocumentApp.getUi().alert("Deleted " + deleted + " unanchored comments.");
  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}
