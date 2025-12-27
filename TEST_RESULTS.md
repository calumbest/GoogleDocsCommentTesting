# Test Results

## Test 1: DOCX Round-Trip Comment Insertion

**Date:** December 26, 2025
**Status:** ✅ PARTIAL SUCCESS

### What We Tested
1. Created Google Doc with test content
2. Exported as .docx
3. Used python-docx to add comment targeting "This is a target phrase for our comment test"
4. Re-uploaded to Google Drive and opened as Google Doc

### Results
- ✅ Comment survived the round-trip
- ✅ Comment is anchored to text (not "original content deleted")
- ⚠️ Comment anchored to **entire paragraph** instead of just the target phrase

### Why Paragraph-Level Anchoring?
The .docx format uses "runs" (`<w:r>` elements) as the atomic unit for comments. When Google Docs exports to .docx, it may split text into multiple runs arbitrarily. Our target phrase likely spanned multiple runs, so the script fell back to commenting the whole paragraph.

### Screenshot
![Test 1 Result](./screenshots/test1_result.png)
- Yellow highlight shows the comment anchor (entire paragraph)
- Comment panel shows "From imported document"

### Next Steps
- **Test 1b:** Improve script to split runs at precise character positions before adding comments
- This should allow targeting specific phrases rather than whole paragraphs

---

## Test 1b: Precise Run Targeting

**Date:** December 26, 2025
**Status:** ✅ SUCCESS

### What We Tested
Used `test_docx_precise.py` which:
1. Analyzes the run structure of the exported .docx
2. Splits runs at character boundaries to isolate target text
3. Attaches comment only to the runs containing the target phrase

### Results
- ✅ Comment anchored to **exactly** the target phrase
- ✅ "This is a target phrase for our comment test" highlighted (without surrounding text)
- ✅ Precision confirmed: period after phrase was NOT included (as intended)

### Key Technical Details
The .docx from Google Docs had the entire paragraph as a single run. The script:
- Found target at character positions 68-112 within the run
- Split the run into 3 parts: before (chars 0-67), target (68-112), after (113-191)
- Attached comment only to the middle run

### Conclusion
**DOCX round-trip with run splitting is a viable solution for precise comment placement in Google Docs.**

---

## Summary: Viable Approach Found

The DOCX round-trip method works:
1. Export Google Doc as .docx
2. Use python-docx to split runs and add comments at precise character positions
3. Re-upload to Google Drive and convert back to Google Docs
4. Comments survive with correct anchoring

### Known Limitations
- **Google account linkage is lost** - See Test 2 below

### Limitations to Explore
- What happens when document content changes after import?
- Performance with large documents?
- Handling comments that span paragraphs?

---

## Test 2: Comment Preservation Round-Trip

**Date:** December 26, 2025
**Status:** ✅ SUCCESS (with limitations)

### What We Tested
1. Created Google Doc with test content
2. Added a manual comment via Google Docs UI (commented on "ipsum")
3. Exported as .docx
4. Used python-docx to add a NEW comment to "quick brown fox"
5. Re-uploaded to Google Drive and opened as Google Doc

### Results
- ✅ Original comment preserved with correct anchor position
- ✅ Original comment text preserved ("Test comment from Calum Best at 3:17pm")
- ✅ Original author name preserved
- ✅ Original timestamp preserved
- ✅ New programmatic comment appeared and anchored correctly
- ✅ Multiple comments coexist without issues

### Important Limitation: Google Account Linkage Lost

⚠️ **All comments show "From imported document"** - Google account linkage is broken.

When a Google Doc is exported to .docx and re-imported:
- Author names are preserved as **plain text strings**
- Profile pictures are lost
- Comments are no longer linked to Google accounts
- Comments appear with "From imported document" label

**Why this happens:** The .docx format stores author as a simple string field, not a Google account identifier. There's no OAuth token or account reference that could survive the export/import cycle.

**Impact:** Users cannot:
- Click on author name to see Google profile
- @ mention the original commenter in replies

**Acceptable for:** Automated/programmatic commenting where account linkage isn't critical.

**Not suitable for:** Workflows where maintaining Google account associations is required.

---

## Test 3: Apps Script / Drive API Approach

**Date:** December 26, 2025
**Status:** ❌ FAILED (confirms known limitation)

### What We Tested
1. Attempted to use `DocumentApp.addComment()` - method doesn't exist
2. Used Drive Advanced Service (`Drive.Comments.create()`) to add comment
3. Listed existing comments to examine anchor format

### Results
- ❌ `DocumentApp` has no comment methods
- ❌ `Drive.Comments.create()` creates **unanchored** comments only
- ✅ Can read existing comment anchors via `Drive.Comments.list()`

### Key Data: Anchor Format Comparison

**Manual comment (created in Google Docs UI):**
```json
{
  "anchor": "kix.hkaqsaj0l7p3",
  "quotedFileContent": {"mimeType": "text/html", "value": "ipsum"}
}
```

**API comment (created via Drive.Comments.create):**
```json
{
  "anchor": undefined,
  "quotedFileContent": undefined
}
```

### Conclusion
The `kix.xxxxx` anchor format is proprietary and generated internally by Google Docs. There is no documented way to construct valid anchors via API. This approach **cannot** create anchored comments.

---

## Test 3b: Pass kix Anchor Directly

**Date:** December 26, 2025
**Status:** ⚠️ PARTIAL SUCCESS

### What We Tested
Passed a captured kix anchor (`kix.hkaqsaj0l7p3`) directly to `Drive.Comments.create()`

### Results
- ✅ Comment created with anchor preserved in API response
- ✅ Comment panel shows "Tab 1" (not "Original content deleted")
- ✅ Clicking comment navigates to correct text
- ❌ No yellow highlight in document body (partial anchoring)

---

## Test 3c: Reuse Existing Anchor

**Date:** December 26, 2025
**Status:** ✅ **BREAKTHROUGH SUCCESS**

### What We Tested
1. Created a manual comment in Google Docs UI (anchored to "ipsum")
2. Read the comment's kix anchor via Drive API
3. Created a NEW comment via API using the SAME anchor

### Results
- ✅ **Full yellow highlighting** on "ipsum" in document body
- ✅ **Google account linkage preserved** (shows user profile, clickable name)
- ✅ Comment appears in normal document interface
- ✅ Anchor returned in API response

### Key Finding
**The Drive API CAN create fully anchored comments if you reuse a valid existing kix anchor!**

This means the workflow is:
1. Create a "template" comment manually (or via DOCX import) to generate a kix anchor
2. Read the anchor via `Drive.Comments.list()`
3. Create new comments via `Drive.Comments.create()` with that anchor
4. New comments inherit full anchoring AND preserve Google account linkage

### Observations
- When the original "owner" comment is deleted, other comments using that anchor may show "original content deleted"
- Only one comment per kix anchor appears in the main document interface at a time
- More testing needed to understand anchor lifecycle

### Screenshot
Test 3c comment showing full anchoring with account linkage preserved.
