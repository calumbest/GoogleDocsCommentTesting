/**
 * Google Docs Comment Tools via Drive API
 *
 * SETUP:
 * 1. Open Google Doc → Extensions → Apps Script
 * 2. Click "+" next to "Services" → Add "Drive API"
 * 3. Paste this code, save
 * 4. Run functions as needed
 */

// ============================================================
// READING COMMENTS
// ============================================================

/**
 * List all comments with their anchors
 * Use this to find kix anchors from DOCX-imported comments
 */
function listAllComments() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  try {
    const comments = Drive.Comments.list(fileId, {
      fields: "comments(id,content,author,anchor,quotedFileContent,deleted)"
    });

    if (!comments.comments || comments.comments.length === 0) {
      Logger.log("No comments found");
      DocumentApp.getUi().alert("No comments found in this document.");
      return;
    }

    Logger.log("=== ALL COMMENTS ===\n");

    comments.comments.forEach(function(c, i) {
      const hasAnchor = c.anchor ? "YES" : "NO";
      const quoted = c.quotedFileContent ? c.quotedFileContent.value : "none";

      Logger.log("Comment " + (i+1) + ":");
      Logger.log("  ID: " + c.id);
      Logger.log("  Content: " + c.content);
      Logger.log("  Author: " + (c.author ? c.author.displayName : "unknown"));
      Logger.log("  Anchor: " + (c.anchor || "NONE"));
      Logger.log("  Quoted text: " + quoted);
      Logger.log("  Deleted: " + c.deleted);
      Logger.log("");
    });

    DocumentApp.getUi().alert(
      "Found " + comments.comments.length + " comments.\n\n" +
      "Check View → Logs for details including kix anchors."
    );

    return comments.comments;
  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}

// ============================================================
// TEST 4: DOCX → API HYBRID WORKFLOW
// ============================================================

/**
 * Step 1: Find the DOCX-imported comment and show its anchor
 */
function test4_step1_findDocxAnchor() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  try {
    const comments = Drive.Comments.list(fileId, {
      fields: "comments(id,content,anchor,quotedFileContent)"
    });

    // Look for comments that have anchors (likely from DOCX import)
    let found = [];
    comments.comments.forEach(function(c) {
      if (c.anchor) {
        found.push(c);
        Logger.log("Found anchored comment:");
        Logger.log("  ID: " + c.id);
        Logger.log("  Content: " + c.content);
        Logger.log("  Anchor: " + c.anchor);
        Logger.log("  Quoted: " + (c.quotedFileContent ? c.quotedFileContent.value : "none"));
      }
    });

    if (found.length === 0) {
      DocumentApp.getUi().alert("No anchored comments found.\n\nMake sure you've uploaded a .docx with comments.");
      return;
    }

    // Store the first anchor for use in step 2
    PropertiesService.getScriptProperties().setProperty('test4_anchor', found[0].anchor);
    PropertiesService.getScriptProperties().setProperty('test4_commentId', found[0].id);
    PropertiesService.getScriptProperties().setProperty('test4_quoted',
      found[0].quotedFileContent ? JSON.stringify(found[0].quotedFileContent) : '');

    DocumentApp.getUi().alert(
      "Found " + found.length + " anchored comment(s).\n\n" +
      "Stored anchor: " + found[0].anchor + "\n" +
      "Comment ID: " + found[0].id + "\n\n" +
      "Now run: test4_step2_deleteAndRecreate()"
    );
  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}

/**
 * Step 2: Delete the DOCX comment, then create a new one with same anchor
 */
function test4_step2_deleteAndRecreate() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  // Get stored values from step 1
  const props = PropertiesService.getScriptProperties();
  const anchor = props.getProperty('test4_anchor');
  const commentId = props.getProperty('test4_commentId');
  const quotedStr = props.getProperty('test4_quoted');

  if (!anchor || !commentId) {
    DocumentApp.getUi().alert("No stored anchor found.\n\nRun test4_step1_findDocxAnchor() first.");
    return;
  }

  Logger.log("=== TEST 4 STEP 2 ===");
  Logger.log("Stored anchor: " + anchor);
  Logger.log("Comment to delete: " + commentId);

  try {
    // Delete the original DOCX-imported comment
    Logger.log("\nDeleting original comment...");
    Drive.Comments.remove(fileId, commentId);
    Logger.log("✓ Deleted comment " + commentId);

    // Small delay to let deletion propagate
    Utilities.sleep(1000);

    // Create new comment with the same anchor
    Logger.log("\nCreating new comment with same anchor...");

    const newComment = {
      content: "TEST 4 SUCCESS: This comment was created via API using a DOCX-generated anchor!",
      anchor: anchor
    };

    // Add quotedFileContent if we have it
    if (quotedStr) {
      newComment.quotedFileContent = JSON.parse(quotedStr);
    }

    const result = Drive.Comments.create(newComment, fileId, {
      fields: "id,content,anchor,quotedFileContent,author"
    });

    Logger.log("\n✓ New comment created:");
    Logger.log("  ID: " + result.id);
    Logger.log("  Anchor: " + result.anchor);
    Logger.log("  Author: " + (result.author ? result.author.displayName : "unknown"));

    DocumentApp.getUi().alert(
      "TEST 4 COMPLETE!\n\n" +
      "1. Deleted DOCX-imported comment\n" +
      "2. Created new API comment with same anchor\n\n" +
      "CHECK THE DOCUMENT:\n" +
      "- Is the new comment anchored (yellow highlight)?\n" +
      "- Does it show YOUR Google account?\n\n" +
      "If both yes → DOCX→API hybrid works!"
    );

    // Clear stored properties
    props.deleteProperty('test4_anchor');
    props.deleteProperty('test4_commentId');
    props.deleteProperty('test4_quoted');

  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}

/**
 * Alternative: Don't delete, just create alongside
 */
function test4_alternative_keepBoth() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  const props = PropertiesService.getScriptProperties();
  const anchor = props.getProperty('test4_anchor');
  const quotedStr = props.getProperty('test4_quoted');

  if (!anchor) {
    DocumentApp.getUi().alert("No stored anchor found.\n\nRun test4_step1_findDocxAnchor() first.");
    return;
  }

  try {
    const newComment = {
      content: "TEST 4-ALT: Created alongside DOCX comment (not replacing it)",
      anchor: anchor
    };

    if (quotedStr) {
      newComment.quotedFileContent = JSON.parse(quotedStr);
    }

    const result = Drive.Comments.create(newComment, fileId, {
      fields: "id,content,anchor,author"
    });

    Logger.log("Created comment alongside original:");
    Logger.log("  ID: " + result.id);
    Logger.log("  Anchor: " + result.anchor);

    DocumentApp.getUi().alert(
      "Created new comment WITHOUT deleting original.\n\n" +
      "Check: Do both comments appear? Which one shows in the document?"
    );
  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}

// ============================================================
// UTILITIES
// ============================================================

/**
 * Delete all comments (use with caution!)
 */
function deleteAllComments() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  const ui = DocumentApp.getUi();
  const response = ui.alert(
    'Delete All Comments',
    'Are you sure you want to delete ALL comments in this document?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    const comments = Drive.Comments.list(fileId, {
      fields: "comments(id)"
    });

    let deleted = 0;
    comments.comments.forEach(function(c) {
      Drive.Comments.remove(fileId, c.id);
      deleted++;
    });

    Logger.log("Deleted " + deleted + " comments");
    ui.alert("Deleted " + deleted + " comments.");
  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    ui.alert("Error: " + e.toString());
  }
}

/**
 * Create a simple unanchored comment (for comparison)
 */
function createUnanchoredComment() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  try {
    const result = Drive.Comments.create(
      { content: "UNANCHORED: This comment has no anchor" },
      fileId,
      { fields: "id,anchor" }
    );

    DocumentApp.getUi().alert(
      "Created unanchored comment.\n\n" +
      "ID: " + result.id + "\n" +
      "Anchor: " + (result.anchor || "NONE")
    );
  } catch (e) {
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}
