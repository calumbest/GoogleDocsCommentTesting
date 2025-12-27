# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Google Commenter is an exploration project to find a workable solution for programmatically inserting comments with specific text locations into Google Docs.

**Status:** Active research. See [TEST_RESULTS.md](./TEST_RESULTS.md) for detailed findings.

## The Core Problem

Google's `kix.xxxxx` anchor format is proprietary and internally generated. There is no public API to:
1. Create new anchors at arbitrary text positions
2. Translate character positions to kix anchors
3. Generate anchors that produce full UI integration

This limitation has been documented since 2016 (Google Issue #36763384) and remains unresolved.

## Major Discovery: Internal API Format (Test 6)

By reverse-engineering network traffic (HAR analysis), we discovered Google Docs' internal format:

```json
// The /save endpoint uses CHARACTER INDICES for anchoring:
{
  "ty": "as",
  "st": "doco_anchor",
  "si": 282,           // START character index
  "ei": 326,           // END character index
  "sm": {"das_a": {"cv": {"op": "insert", "opValue": "kix.xxxxx"}}}
}
```

**Key insight:** The internal API uses `si` (start index) and `ei` (end index) - exactly what the public API lacks!

**Blocker:** Authentication (XSRF tokens, cookies) prevents external scripts from calling these endpoints.

**Current approach:** Chrome Extension (Test 7) to run in authenticated page context.

## Findings Summary

| Approach | Anchored | Account Linkage | Viable? |
|----------|----------|-----------------|---------|
| DOCX round-trip | ✅ Full | ❌ Lost | **Yes** |
| Drive API alone | ❌ None | ✅ Preserved | No |
| API + UI anchor reuse | ✅ Full | ✅ Preserved | **Limited** |
| Internal API (external) | ✅ Full? | ✅ Preserved? | No (auth blocked) |
| Chrome Extension | ✅ Full? | ✅ Preserved? | **Testing** |

## Viable Workflows (So Far)

### Workflow 1: DOCX Round-Trip
For comments at arbitrary positions when account linkage isn't needed.

```
Google Doc → Export .docx → Add comments with python-docx → Re-upload
```

Key script: `test_docx_precise.py`

### Workflow 2: API Anchor Reuse
For multiple comments at the SAME location with account linkage preserved.

```
Create manual comment → Read anchor via Drive API → Create API comments with same anchor
```

Key script: `apps_script_comments.js`

## Key Files

- `TEST_RESULTS.md` - Detailed test results (Tests 1-6 documented)
- `RESEARCH_FINDINGS.md` - Initial research on all approaches
- `test_docx_precise.py` - DOCX comment insertion with run splitting
- `apps_script_comments.js` - Drive API comment tools
- `test_internal_api.py` - Internal API experiment (blocked by auth)
- `network_capture.har` - Captured network traffic showing internal API format
- `archive/` - Old test scripts and files

## Development

- **Python 3.x** with `python-docx`, `requests`, `browser-cookie3`
- **Google Apps Script** with Drive Advanced Service
- **Chrome Extension** (next test)
