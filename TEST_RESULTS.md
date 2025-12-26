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

## Test 1b: Precise Run Targeting (Pending)

**Status:** 🔄 In Progress

Improving the script to:
1. Find the target text
2. Split runs so target text is isolated
3. Attach comment only to the target runs
