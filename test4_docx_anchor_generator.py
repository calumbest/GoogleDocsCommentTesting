#!/usr/bin/env python3
"""
Test 4: DOCX Anchor Generator

Creates a .docx file with a comment at a precise location.
The comment will generate a kix anchor when uploaded to Google Docs.
That anchor can then be reused via Drive API for account-linked comments.

Usage:
    python test4_docx_anchor_generator.py input.docx output.docx "target text" "comment text"

Example:
    python test4_docx_anchor_generator.py test_input.docx test_anchored.docx "quick brown fox" "ANCHOR GENERATOR"
"""

import sys
import copy
from docx import Document
from docx.oxml.ns import qn


def split_run_at_text(paragraph, target_text: str):
    """Find target_text and split runs to isolate it."""
    full_text = ""
    run_boundaries = []

    for i, run in enumerate(paragraph.runs):
        start = len(full_text)
        full_text += run.text
        end = len(full_text)
        run_boundaries.append((start, end, i))

    target_start = full_text.find(target_text)
    if target_start == -1:
        return None

    target_end = target_start + len(target_text)
    runs_to_split = []

    for start, end, run_idx in run_boundaries:
        if start < target_end and end > target_start:
            runs_to_split.append({
                'run_idx': run_idx,
                'run_start': start,
                'run_end': end,
                'target_start_in_run': max(0, target_start - start),
                'target_end_in_run': min(end - start, target_end - start)
            })

    target_runs = []

    for split_info in runs_to_split:
        run_idx = split_info['run_idx']
        run = paragraph.runs[run_idx]
        run_text = run.text

        t_start = split_info['target_start_in_run']
        t_end = split_info['target_end_in_run']

        before_text = run_text[:t_start]
        target_part = run_text[t_start:t_end]
        after_text = run_text[t_end:]

        if before_text == "" and after_text == "":
            target_runs.append(run)
        else:
            run_element = run._element

            if before_text:
                new_run_before = copy.deepcopy(run_element)
                for t in new_run_before.findall(qn('w:t')):
                    t.text = before_text
                run_element.addprevious(new_run_before)

            if after_text:
                new_run_after = copy.deepcopy(run_element)
                for t in new_run_after.findall(qn('w:t')):
                    t.text = after_text
                run_element.addnext(new_run_after)

            for t in run_element.findall(qn('w:t')):
                t.text = target_part

            target_runs.append(run)

    return target_runs


def add_comment(doc: Document, target_text: str, comment_text: str, author: str = "Anchor Generator"):
    """Add a comment to the target text."""
    for paragraph in doc.paragraphs:
        if target_text in paragraph.text:
            target_runs = split_run_at_text(paragraph, target_text)
            if target_runs:
                doc.add_comment(
                    runs=target_runs,
                    text=comment_text,
                    author=author,
                    initials="AG"
                )
                return True
    return False


def main():
    if len(sys.argv) < 5:
        print("Usage: python test4_docx_anchor_generator.py <input.docx> <output.docx> <target_text> <comment_text>")
        print("\nExample:")
        print('  python test4_docx_anchor_generator.py test_input.docx test_anchored.docx "quick brown fox" "ANCHOR GENERATOR"')
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]
    target_text = sys.argv[3]
    comment_text = sys.argv[4]

    print(f"\n=== Test 4: DOCX Anchor Generator ===\n")
    print(f"Input:   {input_file}")
    print(f"Output:  {output_file}")
    print(f"Target:  '{target_text}'")
    print(f"Comment: '{comment_text}'")

    doc = Document(input_file)

    if add_comment(doc, target_text, comment_text):
        doc.save(output_file)
        print(f"\n✓ Comment added to '{target_text}'")
        print(f"✓ Saved: {output_file}")
        print("\n=== Next Steps ===")
        print("1. Upload to Google Drive → Open with Google Docs")
        print("2. Run Apps Script: test_readAnchors() to get the kix anchor")
        print("3. Run Apps Script: test_deleteAndRecreate() to test anchor reuse")
    else:
        print(f"\n✗ Could not find target text: '{target_text}'")
        sys.exit(1)


if __name__ == "__main__":
    main()
