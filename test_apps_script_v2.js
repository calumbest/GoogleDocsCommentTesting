/**
 * Test 3b: Google Apps Script with Drive Advanced Service
 *
 * SETUP:
 * 1. Open your test Google Doc
 * 2. Extensions > Apps Script
 * 3. Click "+" next to "Services" in left sidebar
 * 4. Find "Drive API" and click "Add"
 * 5. Paste this script and save
 * 6. Run "test_addCommentViaDriveAPI"
 */

/**
 * Test: Add comment using Drive API
 * This creates an unanchored comment (known limitation)
 */
function test_addCommentViaDriveAPI() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  Logger.log("Document ID: " + fileId);

  // Create a comment resource
  const comment = {
    content: "DRIVE API TEST: This comment was added via Drive.Comments.create()"
  };

  try {
    // Use Drive Advanced Service to create comment
    const result = Drive.Comments.create(comment, fileId, {
      fields: "id,content,author,createdTime,anchor"
    });

    Logger.log("SUCCESS! Comment created:");
    Logger.log("  ID: " + result.id);
    Logger.log("  Content: " + result.content);
    Logger.log("  Author: " + JSON.stringify(result.author));
    Logger.log("  Anchor: " + result.anchor);

    DocumentApp.getUi().alert(
      "Comment created!\n\n" +
      "ID: " + result.id + "\n" +
      "Anchor: " + (result.anchor || "NONE - unanchored") + "\n\n" +
      "Check if it appears in the document's comment panel."
    );
  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}

/**
 * Test: Try to add anchored comment with custom anchor
 * This tests various anchor formats to see if any work
 */
function test_addAnchoredComment() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  // Try different anchor formats that have been suggested online
  const anchorFormats = [
    // Format 1: Simple text anchor
    JSON.stringify({ r: "head", a: [{ txt: { o: 61, l: 15, ml: true }}]}),

    // Format 2: Line-based anchor
    JSON.stringify({ region: { kind: "drive#commentRegion", line: 1 }}),
  ];

  for (let i = 0; i < anchorFormats.length; i++) {
    const anchor = anchorFormats[i];
    Logger.log("Trying anchor format " + (i+1) + ": " + anchor);

    const comment = {
      content: "ANCHOR TEST " + (i+1) + ": Testing anchor format",
      anchor: anchor
    };

    try {
      const result = Drive.Comments.create(comment, fileId, {
        fields: "id,content,anchor"
      });

      Logger.log("  Result - ID: " + result.id + ", Anchor: " + result.anchor);
    } catch (e) {
      Logger.log("  Failed: " + e.toString());
    }
  }

  DocumentApp.getUi().alert(
    "Attempted to create comments with different anchor formats.\n\n" +
    "Check the Execution Log (View > Logs) for details.\n" +
    "Check the document to see if any comments are anchored."
  );
}

/**
 * List all existing comments to see their anchor format
 */
function test_listExistingComments() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();

  try {
    const comments = Drive.Comments.list(fileId, {
      fields: "comments(id,content,author,anchor,quotedFileContent)"
    });

    if (!comments.comments || comments.comments.length === 0) {
      DocumentApp.getUi().alert("No comments found in this document.");
      return;
    }

    let output = "Found " + comments.comments.length + " comment(s):\n\n";

    comments.comments.forEach(function(c, i) {
      output += "--- Comment " + (i+1) + " ---\n";
      output += "Content: " + c.content + "\n";
      output += "Author: " + (c.author ? c.author.displayName : "unknown") + "\n";
      output += "Anchor: " + (c.anchor || "NONE") + "\n";
      output += "Quoted text: " + (c.quotedFileContent ? c.quotedFileContent.value : "NONE") + "\n\n";

      Logger.log("Comment " + (i+1) + ":");
      Logger.log("  Content: " + c.content);
      Logger.log("  Anchor: " + c.anchor);
      Logger.log("  QuotedFileContent: " + JSON.stringify(c.quotedFileContent));
    });

    DocumentApp.getUi().alert(output);
  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}

/**
 * Utility: Get document info
 */
function test_getDocInfo() {
  const doc = DocumentApp.getActiveDocument();

  const info = {
    id: doc.getId(),
    name: doc.getName(),
    url: doc.getUrl()
  };

  Logger.log(JSON.stringify(info, null, 2));
  DocumentApp.getUi().alert(
    "Document Info:\n\n" +
    "ID: " + info.id + "\n" +
    "Name: " + info.name + "\n" +
    "URL: " + info.url
  );
}
