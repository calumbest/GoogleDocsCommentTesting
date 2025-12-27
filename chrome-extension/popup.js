/**
 * Popup script - communicates with content script
 */

const docInfoEl = document.getElementById('docInfo');
const targetTextEl = document.getElementById('targetText');
const commentTextEl = document.getElementById('commentText');
const addCommentBtn = document.getElementById('addCommentBtn');
const getTextBtn = document.getElementById('getTextBtn');
const statusEl = document.getElementById('status');
const outputEl = document.getElementById('output');

// Get current tab
async function getCurrentTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab;
}

// Send message to content script
async function sendMessage(message) {
  const tab = await getCurrentTab();
  return chrome.tabs.sendMessage(tab.id, message);
}

// Show status
function showStatus(message, type = 'loading') {
  statusEl.textContent = message;
  statusEl.className = 'status ' + type;
}

// Show output
function showOutput(data) {
  outputEl.innerHTML = '<pre>' + JSON.stringify(data, null, 2) + '</pre>';
}

// Load document info on popup open
async function loadDocInfo() {
  try {
    const tab = await getCurrentTab();

    // Check if we're on a Google Docs page
    if (!tab.url.includes('docs.google.com/document')) {
      docInfoEl.innerHTML = '<strong>Not a Google Doc</strong><br>Open a Google Doc to use this extension.';
      addCommentBtn.disabled = true;
      getTextBtn.disabled = true;
      return;
    }

    const info = await sendMessage({ action: 'getDocumentInfo' });

    if (info.error) {
      docInfoEl.innerHTML = '<strong>Error:</strong> ' + info.error;
      return;
    }

    docInfoEl.innerHTML = `
      <strong>Document:</strong> ${info.title}<br>
      <strong>ID:</strong> <code>${info.docId}</code><br>
      <strong>Editor loaded:</strong> ${info.hasEditor ? 'Yes' : 'No'}
    `;
  } catch (e) {
    docInfoEl.innerHTML = `<strong>Error:</strong> ${e.message}<br><em>Make sure you're on a Google Doc page and refresh it.</em>`;
    addCommentBtn.disabled = true;
  }
}

// Get document text
getTextBtn.addEventListener('click', async () => {
  showStatus('Extracting document text...', 'loading');

  try {
    const result = await sendMessage({ action: 'getDocumentText' });

    if (result.error) {
      showStatus('Error: ' + result.error, 'error');
      return;
    }

    showStatus(`Found ${result.text.length} characters (method: ${result.method})`, 'success');
    showOutput({
      method: result.method,
      length: result.text.length,
      preview: result.text.substring(0, 500) + (result.text.length > 500 ? '...' : '')
    });
  } catch (e) {
    showStatus('Error: ' + e.message, 'error');
  }
});

// Add comment
addCommentBtn.addEventListener('click', async () => {
  const targetText = targetTextEl.value.trim();
  const commentText = commentTextEl.value.trim();

  if (!targetText) {
    showStatus('Please enter target text', 'error');
    return;
  }
  if (!commentText) {
    showStatus('Please enter comment text', 'error');
    return;
  }

  showStatus('Creating comment...', 'loading');
  addCommentBtn.disabled = true;

  try {
    const result = await sendMessage({
      action: 'addComment',
      targetText,
      commentText
    });

    if (result.success) {
      showStatus('Comment created! Check the document.', 'success');
      showOutput(result.result);
    } else {
      showStatus('Error: ' + result.error, 'error');
    }
  } catch (e) {
    showStatus('Error: ' + e.message, 'error');
  } finally {
    addCommentBtn.disabled = false;
  }
});

// Initialize
loadDocInfo();
