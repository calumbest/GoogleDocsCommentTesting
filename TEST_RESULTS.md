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
- See the author's profile picture

**Acceptable for:** Automated/programmatic commenting where account linkage isn't critical.

**Not suitable for:** Workflows where maintaining Google account associations is required.
