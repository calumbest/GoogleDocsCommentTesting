/**
 * Test 3: Google Apps Script addComment() with Range
 *
 * This script tests whether Apps Script can create comments
 * that are properly anchored to specific text ranges.
 *
 * SETUP:
 * 1. Open your test Google Doc
 * 2. Extensions > Apps Script
 * 3. Delete any existing code and paste this entire script
 * 4. Save (Ctrl+S)
 * 5. Run > Run function > select "test_addCommentToPhrase"
 * 6. Authorize when prompted
 * 7. Check if comment appears anchored in the document
 */

/**
 * Test 1: Add comment to a specific phrase using findText()
 */
function test_addCommentToPhrase() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  // Search for our target phrase
  const targetPhrase = "quick brown fox";
  const searchResult = body.findText(targetPhrase);

  if (!searchResult) {
    Logger.log("ERROR: Could not find phrase: " + targetPhrase);
    DocumentApp.getUi().alert("Could not find phrase: " + targetPhrase);
    return;
  }

  // Get the element and offsets
  const element = searchResult.getElement();
  const startOffset = searchResult.getStartOffset();
  const endOffset = searchResult.getEndOffsetInclusive();

  Logger.log("Found phrase at offsets: " + startOffset + " to " + endOffset);
  Logger.log("Element type: " + element.getType());
  Logger.log("Element text: " + element.asText().getText());

  // Build a range for just this phrase
  const rangeBuilder = doc.newRange();
  rangeBuilder.addElement(element, startOffset, endOffset);
  const range = rangeBuilder.build();

  // Add the comment
  const commentText = "APPS SCRIPT TEST: This comment was added via Apps Script targeting '" + targetPhrase + "'";

  try {
    doc.addComment(commentText, range);
    Logger.log("SUCCESS: Comment added!");
    DocumentApp.getUi().alert("Comment added! Check if it's anchored to: " + targetPhrase);
  } catch (e) {
    Logger.log("ERROR adding comment: " + e.toString());
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}

/**
 * Test 2: Add comment to entire first paragraph
 */
function test_addCommentToParagraph() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  // Get first non-empty paragraph
  let targetPara = null;
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const text = child.asText().getText().trim();
      if (text.length > 0) {
        targetPara = child;
        break;
      }
    }
  }

  if (!targetPara) {
    DocumentApp.getUi().alert("No paragraph found");
    return;
  }

  Logger.log("Target paragraph: " + targetPara.asText().getText());

  // Build range for entire paragraph
  const rangeBuilder = doc.newRange();
  rangeBuilder.addElement(targetPara);
  const range = rangeBuilder.build();

  try {
    doc.addComment("APPS SCRIPT TEST: Comment on entire paragraph", range);
    DocumentApp.getUi().alert("Comment added to paragraph: " + targetPara.asText().getText().substring(0, 30) + "...");
  } catch (e) {
    DocumentApp.getUi().alert("Error: " + e.toString());
  }
}

/**
 * Test 3: List all comments in the document (read test)
 */
function test_listComments() {
  const doc = DocumentApp.getActiveDocument();

  // Note: DocumentApp doesn't have a direct way to list comments
  // We need to use Drive API for that

  // But we can check if the document has any comment-related metadata
  const docId = doc.getId();
  Logger.log("Document ID: " + docId);
  Logger.log("Document name: " + doc.getName());

  DocumentApp.getUi().alert(
    "Document ID: " + docId + "\n\n" +
    "Note: To list existing comments, you'd need to use the Drive API.\n" +
    "Apps Script's DocumentApp doesn't expose comment listing."
  );
}

/**
 * Utility: Show document structure
 */
function test_showStructure() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  let output = "Document Structure:\n\n";

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    const type = child.getType();
    let preview = "";

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      preview = child.asText().getText().substring(0, 40);
    }

    output += i + ": " + type + " - \"" + preview + "...\"\n";
  }

  Logger.log(output);
  DocumentApp.getUi().alert(output);
}
