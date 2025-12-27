/**
 * Google Docs Comment Injector - Content Script
 *
 * This script runs in the context of Google Docs pages and has access
 * to the authenticated session, allowing us to call internal APIs.
 */

console.log('[Comment Injector] Content script loaded');

// Listen for messages from popup
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  console.log('[Comment Injector] Received message:', request);

  if (request.action === 'getDocumentInfo') {
    getDocumentInfo().then(sendResponse);
    return true; // Keep channel open for async response
  }

  if (request.action === 'addComment') {
    addComment(request.targetText, request.commentText)
      .then(result => sendResponse({ success: true, result }))
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true;
  }

  if (request.action === 'getDocumentText') {
    getDocumentText().then(sendResponse);
    return true;
  }
});

/**
 * Extract document information from the page
 */
async function getDocumentInfo() {
  try {
    // Document ID from URL
    const urlMatch = window.location.href.match(/\/document\/d\/([^/]+)/);
    const docId = urlMatch ? urlMatch[1] : null;

    // Try to get session info from page scripts/globals
    // Google Docs stores some data in window variables
    const info = {
      docId,
      url: window.location.href,
      title: document.title,
      hasEditor: !!document.querySelector('.kix-appview-editor'),
    };

    return info;
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Check if text is "useful" (not just whitespace/control chars)
 */
function isUsefulText(text) {
  if (!text) return false;
  // Remove zero-width spaces, control chars, and whitespace
  const cleaned = text.replace(/[\u200B-\u200D\uFEFF\s]/g, '');
  return cleaned.length >= 5; // At least 5 real characters
}

/**
 * Extract the document text content
 * Google Docs renders in a canvas, but text is also in DOM for accessibility
 */
async function getDocumentText() {
  const results = [];

  try {
    // Method 1: Paragraph renderers (most reliable for newer Google Docs)
    const paragraphs = document.querySelectorAll('.kix-paragraphrenderer');
    if (paragraphs.length > 0) {
      let text = '';
      paragraphs.forEach((p, i) => {
        if (i > 0) text += '\n';
        text += p.textContent;
      });
      if (isUsefulText(text)) {
        results.push({ text, method: 'kix-paragraphrenderer', charCount: text.length });
      }
    }

    // Method 2: Line view content blocks
    const lineBlocks = document.querySelectorAll('.kix-lineview-content');
    if (lineBlocks.length > 0) {
      let text = '';
      lineBlocks.forEach(el => {
        text += el.textContent;
      });
      if (isUsefulText(text)) {
        results.push({ text, method: 'kix-lineview-content', charCount: text.length });
      }
    }

    // Method 3: Word nodes (granular text)
    const wordNodes = document.querySelectorAll('.kix-wordhtmlgenerator-word-node');
    if (wordNodes.length > 0) {
      let text = '';
      wordNodes.forEach(el => {
        text += el.textContent;
      });
      if (isUsefulText(text)) {
        results.push({ text, method: 'kix-wordhtmlgenerator', charCount: text.length });
      }
    }

    // Method 4: Try the page content wrapper
    const pageContent = document.querySelector('.kix-page-content-wrapper');
    if (pageContent) {
      const text = pageContent.innerText;
      if (isUsefulText(text)) {
        results.push({ text, method: 'page-content-wrapper', charCount: text.length });
      }
    }

    // Method 5: Try all spans with specific styling (text spans)
    const allSpans = document.querySelectorAll('.kix-lineview span[style*="font"]');
    if (allSpans.length > 0) {
      let text = '';
      allSpans.forEach(el => {
        text += el.textContent;
      });
      if (isUsefulText(text)) {
        results.push({ text, method: 'styled-spans', charCount: text.length });
      }
    }

    // Method 6: Raw document body scan for any text-containing elements
    const docsBody = document.querySelector('.kix-appview-editor');
    if (docsBody) {
      const walker = document.createTreeWalker(
        docsBody,
        NodeFilter.SHOW_TEXT,
        null,
        false
      );
      let text = '';
      let node;
      while (node = walker.nextNode()) {
        const content = node.textContent.trim();
        if (content && content.length > 0) {
          text += content + ' ';
        }
      }
      if (isUsefulText(text)) {
        results.push({ text: text.trim(), method: 'treewalker', charCount: text.length });
      }
    }

    // Return the result with the most content
    if (results.length > 0) {
      results.sort((a, b) => b.charCount - a.charCount);
      const best = results[0];
      best.allMethods = results.map(r => `${r.method}(${r.charCount})`);
      return best;
    }

    return {
      text: null,
      error: 'Could not extract document text with any method',
      debug: {
        hasParagraphs: document.querySelectorAll('.kix-paragraphrenderer').length,
        hasLineviews: document.querySelectorAll('.kix-lineview').length,
        hasEditor: !!document.querySelector('.kix-appview-editor'),
        hasCanvas: !!document.querySelector('.kix-canvas-tile-content')
      }
    };
  } catch (e) {
    return { text: null, error: e.message };
  }
}

/**
 * Generate a random kix anchor ID
 */
function generateKixAnchor() {
  const chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  let result = 'kix.';
  for (let i = 0; i < 12; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

/**
 * Generate a random comment ID
 */
function generateCommentId() {
  const chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  let result = '';
  for (let i = 0; i < 12; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

/**
 * Extract session parameters from the page
 * These are needed for API calls
 */
async function getSessionParams() {
  // Try to find session data in script tags or globals
  const scripts = document.querySelectorAll('script');
  let sessionData = {
    token: null,
    sid: null,
    ouid: null,
    revision: null
  };

  for (const script of scripts) {
    const content = script.textContent || '';

    // Look for token
    const tokenMatch = content.match(/["']token["']\s*:\s*["']([^"']+)["']/);
    if (tokenMatch) sessionData.token = tokenMatch[1];

    // Look for session ID
    const sidMatch = content.match(/["']sid["']\s*:\s*["']([^"']+)["']/);
    if (sidMatch) sessionData.sid = sidMatch[1];

    // Look for owner/user ID
    const ouidMatch = content.match(/["']ouid["']\s*:\s*["']([^"']+)["']/);
    if (ouidMatch) sessionData.ouid = ouidMatch[1];
  }

  // Alternative: Try to extract from URL parameters of existing requests
  // or from window objects

  return sessionData;
}

/**
 * Main function to add a comment
 */
async function addComment(targetText, commentText) {
  console.log('[Comment Injector] Adding comment...');
  console.log('  Target:', targetText);
  console.log('  Comment:', commentText);

  // Step 1: Get document info
  const docInfo = await getDocumentInfo();
  if (!docInfo.docId) {
    throw new Error('Could not determine document ID');
  }
  console.log('  Document ID:', docInfo.docId);

  // Step 2: Get document text and find target position
  const textResult = await getDocumentText();
  if (!textResult.text) {
    throw new Error('Could not extract document text: ' + textResult.error);
  }
  console.log('  Text extraction method:', textResult.method);
  console.log('  Text length:', textResult.text.length);

  const startIndex = textResult.text.indexOf(targetText);
  if (startIndex === -1) {
    throw new Error(`Target text "${targetText}" not found in document`);
  }
  const endIndex = startIndex + targetText.length;
  console.log('  Target found at:', startIndex, '-', endIndex);

  // Step 3: Get session parameters
  const session = await getSessionParams();
  console.log('  Session params:', session);

  // Step 4: Generate IDs
  const kixAnchor = generateKixAnchor();
  const commentId = generateCommentId();
  const timestamp = Date.now();
  console.log('  Generated anchor:', kixAnchor);
  console.log('  Generated comment ID:', commentId);

  // Step 5: Create the anchor via /save endpoint
  const saveResult = await createAnchor(docInfo.docId, session, startIndex, endIndex, kixAnchor);
  console.log('  Anchor creation result:', saveResult);

  // Step 6: Create the comment via /docos/p/sync endpoint
  const commentResult = await createComment(docInfo.docId, session, kixAnchor, commentId, commentText, targetText, timestamp);
  console.log('  Comment creation result:', commentResult);

  return {
    anchor: kixAnchor,
    commentId,
    startIndex,
    endIndex,
    saveResult,
    commentResult
  };
}

/**
 * Create anchor via /save endpoint
 */
async function createAnchor(docId, session, startIndex, endIndex, kixAnchor) {
  const url = `https://docs.google.com/document/d/${docId}/save`;

  // Build query params
  const params = new URLSearchParams({
    id: docId,
    sid: session.sid || 'unknown',
    vc: '1',
    c: '1',
    w: '1',
    flr: '0',
    smv: '2147483647',
    token: session.token || '',
    ouid: session.ouid || '',
    includes_info_params: 'true',
    cros_files: 'false',
    tab: 't.0'
  });

  // Build request body
  const bundle = {
    commands: [{
      ty: 'as',
      st: 'doco_anchor',
      si: startIndex,
      ei: endIndex,
      sm: {
        das_a: {
          cv: {
            op: 'insert',
            opIndex: 0,
            opValue: kixAnchor
          }
        }
      }
    }],
    sid: session.sid || 'unknown',
    reqId: Math.floor(Math.random() * 100)
  };

  const body = new URLSearchParams({
    rev: String((session.revision || 0) + 1),
    bundles: JSON.stringify([bundle])
  });

  try {
    const response = await fetch(`${url}?${params}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8',
        'X-Same-Domain': '1'
      },
      body: body,
      credentials: 'include'
    });

    const text = await response.text();
    return {
      status: response.status,
      ok: response.ok,
      preview: text.substring(0, 200)
    };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Create comment via /docos/p/sync endpoint
 */
async function createComment(docId, session, kixAnchor, commentId, commentText, quotedText, timestamp) {
  const url = `https://docs.google.com/document/d/${docId}/docos/p/sync`;

  const params = new URLSearchParams({
    id: docId,
    reqid: String(Math.floor(Math.random() * 100)),
    sid: session.sid || 'unknown',
    vc: '1',
    c: '1',
    w: '1',
    flr: '0',
    smv: '2147483647',
    token: session.token || '',
    ouid: session.ouid || '',
    includes_info_params: 'true',
    cros_files: 'false',
    tab: 't.0'
  });

  // Comment payload structure (from HAR analysis)
  const payload = [
    [
      [
        commentId,
        [
          null,
          null,
          ['text/html', commentText],
          ['text/plain', commentText],
          [
            'Extension Test',  // Author name
            null,
            null,  // Profile pic
            session.ouid || '',
            1,
            null,
            null,
            null
          ],
          timestamp,
          timestamp,
          null,
          ['text/plain', quotedText],
          null,
          commentId,
          1
        ],
        timestamp,
        null,
        null,
        null,
        null,
        kixAnchor,
        1
      ]
    ],
    timestamp
  ];

  const body = new URLSearchParams({
    p: JSON.stringify(payload)
  });

  try {
    const response = await fetch(`${url}?${params}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8',
        'X-Same-Domain': '1'
      },
      body: body,
      credentials: 'include'
    });

    const text = await response.text();
    return {
      status: response.status,
      ok: response.ok,
      preview: text.substring(0, 200)
    };
  } catch (e) {
    return { error: e.message };
  }
}

console.log('[Comment Injector] Ready');
