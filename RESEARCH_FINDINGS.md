# Google Docs Comment Positioning: Research Findings

**Date:** December 26, 2025
**Purpose:** Comprehensive exploration of approaches to programmatically insert comments with specific text locations into Google Docs

---

## Executive Summary

Creating programmatically anchored comments in Google Docs remains an **unsolved problem** as of December 2025. Google has not exposed the internal `kix` anchor format through any public API, and the Drive API's anchor field does not work reliably for Google Docs. However, several workarounds and alternative approaches exist with varying levels of viability.

---

## 1. API-Based Approaches (Official Google APIs)

### 1.1 Google Drive API v3 - Comments Resource

**What's Known:**
- The Drive API has a `comments` resource with `create`, `list`, `update`, `delete` methods
- Comments have an `anchor` field defined as a JSON string
- The documented anchor format uses a `region` object:
  ```python
  anchor = {
      'region': {
          'kind': 'drive#commentRegion',
          'line': ANCHOR_LINE,
          'rev': 'head'
      }
  }
  ```
- **Critical Limitation:** This format does NOT work for Google Docs. It's designed for simpler file types (images, PDFs, plain text)
- Comments created via API appear in the document but show "original content deleted" - they're unanchored
- Google Issue Tracker #36763384 (opened 2016) remains unresolved

**What Needs Testing:**
- Custom anchor schemas using `[AppID].[PropertyName]` format
- Whether any combination of region classifiers (`txt`, `line`, `matrix`) works for Docs

**Likelihood of Success:** LOW - This is a known, long-standing limitation

**Sources:**
- [Manage comments and replies | Google Drive](https://developers.google.com/workspace/drive/api/guides/manage-comments)
- [Google Issue Tracker #36763384](https://issuetracker.google.com/issues/36763384)
- [Custom Google Docs comment anchoring schema](https://github.com/lmmx/devnotes/wiki/Custom-Google-Docs-comment-anchoring-schema)

---

### 1.2 Google Docs API - batchUpdate

**What's Known:**
- The Docs API handles document content manipulation (insertText, deleteContent, etc.)
- **Comments are NOT part of the Docs API** - no batchUpdate requests exist for comments
- The API does support reading/writing named ranges with specific text positions (startIndex, endIndex)
- Indexes are UTF-16 code unit based

**What Needs Testing:**
- Nothing to test - comments are explicitly handled by Drive API, not Docs API

**Likelihood of Success:** N/A - Wrong API for comments

**Sources:**
- [Google Docs API batchUpdate](https://developers.google.com/docs/api/reference/rest/v1/documents/batchUpdate)
- [Requests | Google Docs](https://developers.google.com/docs/api/reference/rest/v1/documents/request)

---

### 1.3 Google Docs API - Suggestions

**What's Known:**
- The Docs API can READ suggestions but CANNOT CREATE them programmatically
- Google Issue Tracker #287903901 tracks this feature request
- Suggestions from Word (track changes) DO convert to Google Docs suggestions on import
- SuggestionsViewMode options: SUGGESTIONS_INLINE, DEFAULT_FOR_CURRENT_ACCESS, PREVIEW_SUGGESTIONS_ACCEPTED

**What Needs Testing:**
- Create .docx with track changes via python-docx -> upload as Google Doc -> verify suggestions appear

**Likelihood of Success:** MEDIUM (via .docx workaround)

**Sources:**
- [Work with suggestions | Google Docs](https://developers.google.com/docs/api/how-tos/suggestions)
- [Google Issue Tracker #287903901](https://issuetracker.google.com/issues/287903901)

---

### 1.4 Google Workspace Events API (NEW - Developer Preview)

**What's Known:**
- As of July 2025, comment events are available in Developer Preview
- Supported events: user posts comment, user replies to comment
- Requires Cloud Pub/Sub subscription
- Currently READ-ONLY (event notifications, not creation)

**What Needs Testing:**
- Whether event payloads include anchor information that could be reverse-engineered

**Likelihood of Success:** LOW for creation, potentially useful for monitoring

**Sources:**
- [Subscribe to Google Drive events](https://developers.google.com/workspace/events/guides/events-drive)
- [Google Workspace Updates: Drive Events API](https://workspaceupdates.googleblog.com/2025/07/google-drive-events-api-now-available.html)

---

## 2. File Format Approaches

### 2.1 DOCX Export/Import with Comments

**What's Known:**
- Google Docs can export to .docx format (File > Download > Microsoft Word)
- Comments ARE preserved during export with proper positioning
- Google Docs can import .docx files with comments preserved
- The .docx format stores comments in `word/comments.xml` with anchors via `commentRangeStart`/`commentRangeEnd` XML elements

**Workflow:**
1. Export Google Doc as .docx
2. Parse/modify .docx using python-docx or similar
3. Add comments with proper anchoring (requires run-level precision)
4. Re-upload to Google Drive with conversion to Google Docs format

**What Needs Testing:**
- Full round-trip: GDoc -> .docx -> add comments -> GDoc
- Whether comment positions survive the conversion accurately
- Whether there's drift/offset in comment positions

**Likelihood of Success:** HIGH - Most promising approach

**Sources:**
- [Working with Comments | python-docx](https://python-docx.readthedocs.io/en/latest/user/comments.html)
- [How to Insert a comment into a word processing document | Microsoft Learn](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-insert-a-comment-into-a-word-processing-document)

---

### 2.2 Office Open XML (OOXML) Direct Manipulation

**What's Known:**
- .docx files are ZIP archives containing XML files
- Comments stored in `word/comments.xml`
- Comment positions marked by `<w:commentRangeStart>` and `<w:commentRangeEnd>` elements with matching IDs
- `<w:commentReference>` links comments to positions
- Structure is well-documented in MS-DOCX specification

**Example XML structure:**
```xml
<w:p>
  <w:commentRangeStart w:id="0" />
  <w:r>
    <w:t>commented text</w:t>
  </w:r>
  <w:commentRangeEnd w:id="0" />
  <w:r>
    <w:commentReference w:id="0" />
  </w:r>
</w:p>
```

**What Needs Testing:**
- Direct XML manipulation vs. using python-docx abstraction
- Complex scenarios: comments spanning paragraphs, nested structures

**Likelihood of Success:** HIGH

**Sources:**
- [CommentRangeStart Class | Microsoft Learn](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.commentrangestart)
- [MS-DOCX Specification](https://interoperability.blob.core.windows.net/files/MS-DOCX/%5BMS-DOCX%5D.pdf)

---

### 2.3 HTML Export Analysis

**What's Known:**
- Google Docs can export as HTML (Web Page .html zipped)
- Comments ARE included in HTML export as annotations/notes
- Format: letters mark comment locations, full text at document end
- HTML is messy with lots of unnecessary markup

**What Needs Testing:**
- Whether HTML comment markup can be parsed to extract position information
- Whether modified HTML with comments can be imported back

**Likelihood of Success:** LOW - Import back to GDocs is problematic

**Sources:**
- [Convert Google Docs to HTML](https://www.groovypost.com/howto/convert-google-docs-to-html/)

---

## 3. Google Apps Script Approaches

### 3.1 DocumentApp.addComment()

**What's Known:**
- Apps Script has `document.addComment(text, range)` method
- Can create Range objects using `RangeBuilder`
- Range must be built from document elements (paragraphs, runs)

**Example:**
```javascript
function addCommentToText() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const range = doc.newRange()
    .addElement(body.getChild(0))
    .build();
  doc.addComment('Comment text', range);
}
```

**What Needs Testing:**
- Whether comments created this way are actually anchored in the UI
- Position accuracy when targeting specific character ranges
- Whether `editAsText().getRange()` produces usable anchors

**Likelihood of Success:** MEDIUM - Requires testing, mixed reports from developers

**Sources:**
- [Class Document | Apps Script](https://developers.google.com/apps-script/reference/document/document)
- [Class RangeBuilder | Apps Script](https://developers.google.com/apps-script/reference/document/range-builder)

---

### 3.2 Drive Advanced Service in Apps Script

**What's Known:**
- Apps Script can enable Drive API as an advanced service
- Same limitations as Drive API v3 apply
- No special privileges for anchor creation

**Likelihood of Success:** LOW

---

## 4. Third-Party Tools and Libraries

### 4.1 python-docx (v1.2.0+)

**What's Known:**
- Full support for reading and writing comments
- `Document.add_comment(runs, text, author, initials)` method
- Comments must anchor to run boundaries (may require splitting runs)
- Can create complex comment content (formatted text, images)

**Key Code:**
```python
from docx import Document

doc = Document()
para = doc.add_paragraph("Hello, world!")
comment = doc.add_comment(
    runs=para.runs,
    text="This is a comment",
    author="Author Name",
    initials="AN"
)
doc.save("output.docx")
```

**Limitations:**
- Comments must anchor at run boundaries
- May need to split runs for arbitrary character positions

**Likelihood of Success:** HIGH for .docx manipulation

**Sources:**
- [Working with Comments | python-docx](https://python-docx.readthedocs.io/en/latest/user/comments.html)
- [Document objects | python-docx](https://python-docx.readthedocs.io/en/latest/api/document.html)

---

### 4.2 Aspose.Words

**What's Known:**
- Commercial library for .NET/Python/Java/C++
- Full comment support with `CommentRangeStart` and `CommentRangeEnd`
- Can convert between many formats including DOCX
- Does NOT directly support Google Docs format

**What Needs Testing:**
- Comment preservation in DOCX that gets uploaded to Google Docs

**Likelihood of Success:** HIGH for .docx, requires conversion step for GDocs

**Sources:**
- [Working with Comments | Aspose.Words](https://docs.aspose.com/words/python-net/working-with-comments/)
- [Aspose.Words for Python](https://pypi.org/project/aspose-words/)

---

### 4.3 docx4j (Java)

**What's Known:**
- Open source Java library for Office Open XML
- Full document manipulation including track changes
- JAXB-based object representation

**Likelihood of Success:** HIGH for .docx manipulation

**Sources:**
- [docx4j](https://www.docx4java.org/trac/docx4j)

---

### 4.4 google-docs-mcp

**What's Known:**
- MCP server for Google Docs integration with AI assistants
- Has `addComment` tool but with the same limitation: "Comments created via the addComment tool appear in the 'All Comments' list but are not visibly anchored to specific text"

**Likelihood of Success:** LOW - same underlying API limitation

**Sources:**
- [google-docs-mcp | GitHub](https://github.com/a-bonus/google-docs-mcp)

---

## 5. Creative/Experimental Approaches

### 5.1 Browser Automation (Puppeteer/Playwright)

**What's Known:**
- Could theoretically automate the UI workflow of adding comments
- Google Docs uses a custom canvas-based rendering engine ("Kix")
- Text selection and comment insertion would require complex DOM manipulation
- Google actively prevents automation and could detect/block it

**What Needs Testing:**
- Feasibility of selecting text in Google Docs canvas
- Triggering comment dialog programmatically
- Detection/blocking mechanisms

**Likelihood of Success:** LOW - Fragile, possibly against ToS, complex implementation

---

### 5.2 Reverse Engineering the Kix Anchor Format

**What's Known:**
- Google Docs uses internal `kix.XXXXXXX` anchor IDs
- These IDs are visible when listing comments via Drive API
- The format is: `anchor=kix.[random_string]`
- Previous reverse engineering attempts have failed to create valid anchors
- One researcher developed an eye ulcer studying minified Kix JavaScript

**What Needs Testing:**
- Analyzing anchor IDs from multiple comments to find patterns
- Whether the ID encodes position information or is purely random
- Clipboard data manipulation when pasting commented text

**Likelihood of Success:** VERY LOW - Google intentionally keeps this proprietary

**Sources:**
- [How I reverse-engineered Google Docs](https://features.jsomers.net/how-i-reverse-engineered-google-docs/)
- [kix-standalone | GitHub](https://github.com/benjamn/kix-standalone)

---

### 5.3 Named Ranges as Comment Proxies

**What's Known:**
- Google Docs API supports named ranges with precise character positions
- Named ranges auto-update when document content changes
- Could potentially be used to mark locations, then manually add comments

**Workflow:**
1. Create named ranges at desired comment positions via Docs API
2. Have user/script navigate to each named range
3. Manually trigger comment creation

**What Needs Testing:**
- Whether named range positions can be used to guide comment placement
- Integration with Apps Script for semi-automated workflow

**Likelihood of Success:** MEDIUM - Not fully automated but could assist

**Sources:**
- [Work with named ranges | Google Docs](https://developers.google.com/docs/api/how-tos/named-ranges)

---

### 5.4 Track Changes via DOCX Workaround

**What's Known:**
- Word track changes convert to Google Docs suggestions on import
- Suggestions appear anchored to specific text positions
- Could use suggestions as a comment-like mechanism

**Workflow:**
1. Create .docx with track changes at desired positions
2. Upload to Google Docs
3. Track changes become suggestions
4. Optionally convert suggestions to comments manually

**Likelihood of Success:** MEDIUM-HIGH - Good for suggestion-based workflow

---

## 6. Recommended Testing Priority

Based on the research, here is the recommended order for testing approaches:

### Priority 1 (Most Promising)
1. **DOCX Round-Trip** - Export GDoc, add comments with python-docx, re-import
2. **Apps Script addComment()** - Test if Range-based comments actually anchor correctly

### Priority 2 (Viable Alternatives)
3. **Track Changes Workaround** - Create .docx with track changes, import as suggestions
4. **Named Ranges + Manual Comments** - Semi-automated workflow

### Priority 3 (Research/Exploration)
5. **Kix Anchor Analysis** - Study existing anchor patterns for any exploitable structure
6. **Google Workspace Events API** - Monitor comment events for anchor data

### Not Recommended
- Drive API anchor field for GDocs (confirmed broken)
- Browser automation (fragile, ToS concerns)
- Waiting for Google to fix (Issue open since 2016)

---

## 7. Key Open Questions

1. Does the DOCX round-trip preserve exact comment positions, or is there drift?
2. What happens to comments when document content changes after import?
3. Can Apps Script's `addComment()` with Range actually create anchored comments?
4. Is there any pattern to the `kix.XXXXXXX` anchor IDs that could be exploited?
5. Will Google ever expose comment anchor creation in their APIs?

---

## 8. Sources Summary

### Official Google Documentation
- [Manage comments and replies | Google Drive](https://developers.google.com/workspace/drive/api/guides/manage-comments)
- [REST Resource: comments | Google Drive](https://developers.google.com/drive/api/reference/rest/v3/comments)
- [Work with named ranges | Google Docs](https://developers.google.com/workspace/docs/api/how-tos/named-ranges)
- [Work with suggestions | Google Docs](https://developers.google.com/docs/api/how-tos/suggestions)
- [Google Docs API release notes](https://developers.google.com/workspace/docs/release-notes)
- [Class Document | Apps Script](https://developers.google.com/apps-script/reference/document/document)

### Google Issue Tracker
- [Issue #36763384: Provide ability to create a Drive API Comment anchor](https://issuetracker.google.com/issues/36763384)
- [Issue #287903901: Google Docs API Support for Suggested Edits](https://issuetracker.google.com/issues/287903901)
- [Issue #292610078: The anchor property of a comment created using the Drive API](https://issuetracker.google.com/issues/292610078)
- [Issue #357985444: Anchored comments with Drive API not working](https://issuetracker.google.com/issues/357985444)

### Third-Party Documentation
- [Working with Comments | python-docx](https://python-docx.readthedocs.io/en/latest/user/comments.html)
- [How to Insert a comment into a word processing document | Microsoft Learn](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-insert-a-comment-into-a-word-processing-document)
- [Working with Comments | Aspose.Words](https://docs.aspose.com/words/python-net/working-with-comments/)

### Community Discussions
- [Creating anchored comments in Google Docs | Google Groups](https://groups.google.com/g/google-apps-script-community/c/OziLHQJz_K8)
- [Custom Google Docs comment anchoring schema | GitHub](https://github.com/lmmx/devnotes/wiki/Custom-Google-Docs-comment-anchoring-schema)
- [How I reverse-engineered Google Docs | James Somers](https://features.jsomers.net/how-i-reverse-engineered-google-docs/)
