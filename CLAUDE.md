# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Google Commenter was an exploration project to find a workable solution for programmatically inserting comments with specific text locations into Google Docs.

**Status:** Research complete. See [TEST_RESULTS.md](./TEST_RESULTS.md) for full findings.

## The Core Problem

Google's `kix.xxxxx` anchor format is proprietary and internally generated. There is no public API to:
1. Create new anchors at arbitrary text positions
2. Translate character positions to kix anchors
3. Generate anchors that produce full UI integration

This limitation has been documented since 2016 (Google Issue #36763384) and remains unresolved.

## Findings Summary

| Approach | Anchored | Account Linkage | Viable? |
|----------|----------|-----------------|---------|
| DOCX round-trip | ✅ Full | ❌ Lost | **Yes** |
| Drive API alone | ❌ None | ✅ Preserved | No |
| API + UI anchor reuse | ✅ Full | ✅ Preserved | **Limited** |
| API + DOCX anchor reuse | ⚠️ Partial | ✅ Preserved | No |
| Browser automation | N/A | N/A | No |

## Viable Workflows

### Workflow 1: DOCX Round-Trip (Recommended)
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

## Development

- **Python 3.x** with `python-docx` for .docx manipulation
- **Google Apps Script** with Drive Advanced Service for API testing
- **Playwright** (attempted, blocked by Google security)

## Key Files

- `TEST_RESULTS.md` - Detailed test results and conclusions
- `RESEARCH_FINDINGS.md` - Initial research on all approaches
- `test_docx_precise.py` - DOCX comment insertion with run splitting
- `apps_script_comments.js` - Drive API comment tools
- `archive/` - Old test scripts and files
