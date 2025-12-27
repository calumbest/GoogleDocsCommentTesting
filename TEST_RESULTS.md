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

---

## Test 4: DOCX → API Hybrid Workflow

**Date:** December 26, 2025
**Status:** ⚠️ PARTIAL SUCCESS (not fully usable)

### What We Tested
1. Created comment via python-docx at specific location ("quick brown fox")
2. Uploaded .docx to Google Docs (comment appears anchored, but "from imported document")
3. Read the DOCX-generated anchor via Drive API
4. Created new comment via API using that anchor

### Results
- ✅ DOCX-imported comment has anchor (`kix.cmt0` format)
- ✅ API can create comment with that anchor
- ✅ New comment shows user's Google account
- ❌ **No yellow highlighting** - only shows in comment panel
- ❌ Cannot delete DOCX-imported comments via API (permission error)

### Key Finding: Anchor Format Matters

| Anchor Source | Format | Full UI Integration |
|---------------|--------|---------------------|
| Manual UI comment | `kix.hkaqsaj0l7p3` | ✅ Yes |
| DOCX import | `kix.cmt0` | ❌ No (panel only) |

DOCX-generated anchors are "weaker" and don't produce full visual integration when reused.

---

## Test 5: Browser Automation

**Date:** December 26, 2025
**Status:** ❌ FAILED

### What We Tested
Attempted to use Playwright to automate the Google Docs UI:
1. Open document in browser
2. Use Cmd+F to find text
3. Use Cmd+Option+M to add comment

### Results
- ❌ **Google blocks sign-in from automated browsers** (security detection)
- Even if login worked, canvas-based rendering makes DOM automation unreliable

### Conclusion
Browser automation is not a viable approach due to security restrictions and technical complexity.

---

## Test 6: Internal API Reverse Engineering (HAR Analysis)

**Date:** December 26, 2025
**Status:** 🔬 RESEARCH SUCCESS / ⚠️ IMPLEMENTATION BLOCKED

### What We Discovered

By capturing network traffic (HAR file) when manually adding a comment, we reverse-engineered Google Docs' internal API format.

#### The `/save` Endpoint (Creates Anchor)
```json
POST https://docs.google.com/document/d/{docId}/save

{
  "rev": "45",
  "bundles": [{
    "commands": [{
      "ty": "as",
      "st": "doco_anchor",
      "si": 282,           // START CHARACTER INDEX
      "ei": 326,           // END CHARACTER INDEX
      "sm": {
        "das_a": {
          "cv": {
            "op": "insert",
            "opIndex": 0,
            "opValue": "kix.cg51nha9yce6"  // GENERATED ANCHOR
          }
        }
      }
    }],
    "sid": "session_id",
    "reqId": 11
  }]
}
```

#### The `/docos/p/sync` Endpoint (Creates Comment)
```json
POST https://docs.google.com/document/d/{docId}/docos/p/sync

[
  [[
    "comment_id",
    [null, null,
      ["text/html", "Comment text"],
      ["text/plain", "Comment text"],
      ["Author Name", null, "profile_pic_url", "user_id", ...],
      timestamp, timestamp, null,
      ["text/plain", "Quoted/highlighted text"],
      null, "comment_id", 1
    ],
    timestamp, null, null, null, null,
    "kix.anchor_id",  // Links to anchor created in /save
    1
  ]],
  timestamp
]
```

### Key Finding: Character Index Anchoring

The internal API uses **character indices** (`si` and `ei`) to define anchor positions - exactly what the public API lacks!

- `si` = start index (character position from document start)
- `ei` = end index
- The kix anchor is generated client-side and associated with these positions

### Why Implementation Failed

1. **XSRF Token Expiration** - Tokens from HAR file expired quickly
2. **Cookie Insufficiency** - browser_cookie3 couldn't capture all necessary session data
3. **403 Forbidden** - Direct requests to fetch fresh tokens were blocked

### Next Step: Chrome Extension

A Chrome Extension running in the authenticated page context should have:
- Valid XSRF tokens (from the page)
- Full cookie access
- Ability to call internal APIs directly

**Status:** Test 7 pending

---

# Final Conclusions

## Summary of All Approaches Tested

| Approach | Anchored | Account Linkage | Viable? |
|----------|----------|-----------------|---------|
| DOCX round-trip | ✅ Full | ❌ Lost | **Yes** (if account linkage not needed) |
| Drive API alone | ❌ None | ✅ Preserved | No |
| API + UI anchor reuse | ✅ Full | ✅ Preserved | **Limited** (same location only) |
| API + DOCX anchor reuse | ⚠️ Partial | ✅ Preserved | No |
| Browser automation (Playwright) | N/A | N/A | No (blocked by security) |
| Internal API (external script) | ✅ Full? | ✅ Preserved? | No (auth blocked) |
| Chrome Extension (Test 7) | ✅ Full? | ✅ Preserved? | **Testing** |

## Viable Workflows

### Workflow 1: DOCX Round-Trip (Recommended)
**Use when:** You need comments at arbitrary positions and don't need account linkage.

```
Google Doc → Export .docx → Add comments with python-docx → Re-upload → Convert to Google Doc
```

**Pros:**
- Precise positioning at any text location
- Preserves existing comments
- Works reliably

**Cons:**
- All comments show "From imported document"
- No Google account linkage
- Requires export/import cycle

### Workflow 2: API Anchor Reuse
**Use when:** You need multiple comments at the SAME location with account linkage.

```
Create manual comment → Read anchor via API → Create additional API comments with same anchor
```

**Pros:**
- Preserves Google account linkage
- Full UI integration (yellow highlighting)

**Cons:**
- Only works for locations with existing manual comments
- Cannot create anchors at arbitrary positions

## The Fundamental Limitation

Google's `kix.xxxxx` anchor format is proprietary and internally generated. There is no public API to:
1. Create new anchors at arbitrary positions
2. Translate character positions to kix anchors
3. Generate anchors that produce full UI integration

This limitation has been documented since 2016 (Google Issue #36763384) and remains unresolved.

## Recommendations

1. **For programmatic commenting at arbitrary positions:** Use DOCX round-trip
2. **For preserving account linkage:** Create manual comments first, then add replies via API
3. **For exporting comments to Markdown:** Use DOCX export and parse comment positions from XML
4. **For Google to fix this:** Star issue #36763384 on Google Issue Tracker
