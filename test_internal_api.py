#!/usr/bin/env python3
"""
Test 6: Internal Google Docs API

Attempts to use the internal /save and /docos/p/sync endpoints
to create anchored comments, based on HAR file analysis.

WARNING: This is experimental and may not work. It relies on
undocumented internal APIs that could change at any time.

Usage:
    python test_internal_api.py network_capture.har "target text" "comment text"
"""

import json
import sys
import re
import random
import string
import time
from urllib.parse import unquote, urlencode, quote
import requests


def generate_kix_anchor():
    """Generate a random kix anchor ID."""
    # Format appears to be: kix. + 12 lowercase alphanumeric chars
    chars = string.ascii_lowercase + string.digits
    random_part = ''.join(random.choice(chars) for _ in range(12))
    return f"kix.{random_part}"


def generate_comment_id():
    """Generate a random comment ID."""
    # Format appears to be: 12 lowercase alphanumeric chars
    chars = string.ascii_lowercase + string.digits
    return ''.join(random.choice(chars) for _ in range(12))


def extract_session_from_har(har_path):
    """Extract authentication details from HAR file."""
    with open(har_path, 'r') as f:
        har = json.load(f)

    session_data = {
        'cookies': {},
        'doc_id': None,
        'sid': None,
        'token': None,
        'ouid': None,
        'revision': None,
    }

    for entry in har['log']['entries']:
        req = entry['request']

        # Find a /save request to extract params
        if '/save?' in req['url'] and req['method'] == 'POST':
            # Extract query params
            for q in req.get('queryString', []):
                if q['name'] == 'id':
                    session_data['doc_id'] = q['value']
                elif q['name'] == 'sid':
                    session_data['sid'] = q['value']
                elif q['name'] == 'token':
                    session_data['token'] = q['value']
                elif q['name'] == 'ouid':
                    session_data['ouid'] = q['value']

            # Extract cookies
            for cookie in req.get('cookies', []):
                session_data['cookies'][cookie['name']] = cookie['value']

            # Extract revision from POST body
            if 'postData' in req:
                text = unquote(req['postData'].get('text', ''))
                match = re.search(r'rev=(\d+)', text)
                if match:
                    session_data['revision'] = int(match.group(1))

            break

    return session_data


def get_document_text(session_data):
    """
    Fetch the document to get current text and revision.
    This is a simplified attempt - may not work.
    """
    # We'd need to parse the document's actual content
    # For now, we'll rely on the HAR data
    return None


def create_anchored_comment(session_data, start_index, end_index, quoted_text, comment_text):
    """
    Attempt to create an anchored comment using internal API.
    """
    doc_id = session_data['doc_id']
    cookies = session_data['cookies']

    # Generate IDs
    kix_anchor = generate_kix_anchor()
    comment_id = generate_comment_id()
    timestamp = int(time.time() * 1000)

    # Increment revision (this is a guess - may need current rev from server)
    new_rev = session_data['revision'] + 1

    print(f"\n=== Attempting Internal API Comment ===")
    print(f"Document ID: {doc_id}")
    print(f"Target indices: {start_index} - {end_index}")
    print(f"Generated anchor: {kix_anchor}")
    print(f"Comment ID: {comment_id}")
    print(f"Revision: {session_data['revision']} -> {new_rev}")
    print()

    # Step 1: Create the anchor via /save
    save_url = f"https://docs.google.com/document/d/{doc_id}/save"
    save_params = {
        'id': doc_id,
        'sid': session_data['sid'],
        'vc': '1',
        'c': '1',
        'w': '1',
        'flr': '0',
        'smv': '2147483647',
        'token': session_data['token'],
        'ouid': session_data['ouid'],
        'includes_info_params': 'true',
        'cros_files': 'false',
        'tab': 't.0',
    }

    save_body = {
        'rev': str(new_rev),
        'bundles': json.dumps([{
            "commands": [{
                "ty": "as",
                "st": "doco_anchor",
                "si": start_index,
                "ei": end_index,
                "sm": {
                    "das_a": {
                        "cv": {
                            "op": "insert",
                            "opIndex": 0,
                            "opValue": kix_anchor
                        }
                    }
                }
            }],
            "sid": session_data['sid'],
            "reqId": random.randint(1, 100)
        }])
    }

    headers = {
        'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8',
        'Origin': 'https://docs.google.com',
        'X-Same-Domain': '1',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
    }

    print("Step 1: Sending /save request to create anchor...")
    print(f"  URL: {save_url}")
    print(f"  Body: {save_body}")

    try:
        response = requests.post(
            save_url,
            params=save_params,
            data=save_body,
            headers=headers,
            cookies=cookies,
            timeout=30
        )
        print(f"  Response status: {response.status_code}")
        print(f"  Response preview: {response.text[:500]}...")

        if response.status_code != 200:
            print("  FAILED - anchor creation unsuccessful")
            return False

    except Exception as e:
        print(f"  ERROR: {e}")
        return False

    # Step 2: Create the comment via /docos/p/sync
    sync_url = f"https://docs.google.com/document/d/{doc_id}/docos/p/sync"
    sync_params = {
        'id': doc_id,
        'reqid': str(random.randint(1, 100)),
        'sid': session_data['sid'],
        'vc': '1',
        'c': '1',
        'w': '1',
        'flr': '0',
        'smv': '2147483647',
        'token': session_data['token'],
        'ouid': session_data['ouid'],
        'includes_info_params': 'true',
        'cros_files': 'false',
        'tab': 't.0',
    }

    # Comment payload structure (based on HAR analysis)
    comment_payload = [
        [
            [
                comment_id,
                [
                    None,
                    None,
                    ["text/html", comment_text],
                    ["text/plain", comment_text],
                    [
                        "API Test",  # Author name
                        None,
                        None,  # Profile pic
                        session_data['ouid'],
                        1,
                        None,
                        None,
                        None  # Email
                    ],
                    timestamp,
                    timestamp,
                    None,
                    ["text/plain", quoted_text],
                    None,
                    comment_id,
                    1
                ],
                timestamp,
                None,
                None,
                None,
                None,
                kix_anchor,
                1
            ]
        ],
        timestamp
    ]

    sync_body = {'p': json.dumps(comment_payload)}

    print("\nStep 2: Sending /docos/p/sync request to create comment...")
    print(f"  URL: {sync_url}")

    try:
        response = requests.post(
            sync_url,
            params=sync_params,
            data=sync_body,
            headers=headers,
            cookies=cookies,
            timeout=30
        )
        print(f"  Response status: {response.status_code}")
        print(f"  Response preview: {response.text[:500]}...")

        if response.status_code == 200:
            print("\n✓ Requests completed! Check the document.")
            return True
        else:
            print("\n✗ Comment sync failed")
            return False

    except Exception as e:
        print(f"  ERROR: {e}")
        return False


def main():
    if len(sys.argv) < 4:
        print("Usage: python test_internal_api.py <har_file> <start_index> <end_index> <quoted_text> <comment_text>")
        print()
        print("Example:")
        print('  python test_internal_api.py network_capture.har 282 326 "target phrase" "My comment"')
        print()
        print("Note: You need to know the character indices of your target text.")
        print("These can be found by counting characters in the document.")
        sys.exit(1)

    har_file = sys.argv[1]
    start_index = int(sys.argv[2])
    end_index = int(sys.argv[3])
    quoted_text = sys.argv[4]
    comment_text = sys.argv[5]

    print("Extracting session data from HAR file...")
    session_data = extract_session_from_har(har_file)

    if not session_data['doc_id']:
        print("ERROR: Could not extract document ID from HAR file")
        sys.exit(1)

    print(f"  Document ID: {session_data['doc_id']}")
    print(f"  Session ID: {session_data['sid']}")
    print(f"  Revision: {session_data['revision']}")
    print(f"  Cookies found: {len(session_data['cookies'])}")

    success = create_anchored_comment(
        session_data,
        start_index,
        end_index,
        quoted_text,
        comment_text
    )

    if success:
        print("\n=== Check the Google Doc to see if comment appeared! ===")
    else:
        print("\n=== Request failed - see errors above ===")


if __name__ == "__main__":
    main()
