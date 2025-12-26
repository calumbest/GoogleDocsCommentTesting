# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Google Commenter is an exploration project to find a workable solution for programmatically inserting comments with specific text locations into Google Docs.

### The Core Problem

Google's Comments API returns comment locations as uninterpretable strings (the `quotedFileContent.value` or anchor fields). There's no documented way to:
1. Parse these location identifiers to understand where a comment is positioned
2. Construct valid location identifiers to post comments at specific positions

This blocks use cases like:
- Exporting Google Docs to Markdown while preserving comment thread positions
- Programmatically posting comments anchored to specific text ranges

### Known Information

- Exporting GDocs to .docx retains comments with positions (potential XML structure to investigate)
- Comments can be posted via API but without location anchoring
- Last checked: Feb 2025. APIs may have been updated since then.

## Project Approach

This is an **exploration/research project**, not a product build. Goals:
1. Identify the full list of potential approaches
2. Test each approach with short feedback loops
3. Find a workable solution (if one exists)

### Constraints

- No building Google Docs alternatives
- No janky workarounds (bookmarks, etc.) without explicit user approval
- Bite-sized code experiments
- Manual GUI testing may be required - always specify exactly what data/screenshots are needed before requesting manual tests

## Testing Plan

See [RESEARCH_FINDINGS.md](./RESEARCH_FINDINGS.md) for comprehensive research on all approaches.

### Test Priority Order

1. **Test 1: DOCX Round-Trip** (HIGH likelihood)
   - Export GDoc as .docx → add comments with python-docx → re-import
   - Script: `test_docx_roundtrip.py`

2. **Test 2: Apps Script addComment()** (MEDIUM likelihood)
   - Test if Range-based comments actually anchor in UI

3. **Test 3: Track Changes → Suggestions** (MEDIUM-HIGH likelihood)
   - Create .docx with track changes → import → verify suggestions are anchored

4. **Test 4: Named Ranges** (MEDIUM likelihood)
   - Semi-automated workflow with position markers

### Test Results

Results will be documented in `TEST_RESULTS.md` as we complete each test.

## Development

Python 3.x with python-docx for .docx manipulation experiments.
