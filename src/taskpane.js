/**
 * Gemini Copilot for PowerPoint
 * Powered by Google Gemini API from Google AI Studio
 */

// Application State
const state = {
  apiKey: localStorage.getItem('gemini_api_key') || '',
  model: localStorage.getItem('gemini_model') || 'gemini-3.8-flash',
  customInstructions: localStorage.getItem('gemini_instructions') || '',
  isOfficeReady: false,
  isPowerPoint: false,
  currentSlideInfo: {
    index: 1,
    title: '',
    text: '',
    shapeCount: 0
  },
  chatHistory: [],
  lastGeneratedSlide: null
};

// ============================================================================
// Initialization & Office.js Readiness
// ============================================================================

window.addEventListener('DOMContentLoaded', () => {
  initDOMListeners();
  loadSavedSettings();
  updateUI();

  if (typeof Office !== 'undefined') {
    Office.onReady((info) => {
      state.isOfficeReady = true;
      if (info.host === Office.HostType.PowerPoint) {
        state.isPowerPoint = true;
        setConnectionStatus(true, 'Connected to PowerPoint');
        scanActiveSlide();
      } else {
        setConnectionStatus(false, 'Running in standalone browser (Preview mode)');
        simulateBrowserContext();
      }
    });
  } else {
    setConnectionStatus(false, 'Office.js not loaded (Preview mode)');
    simulateBrowserContext();
  }
});

function setConnectionStatus(isConnected, message) {
  const statusEl = document.getElementById('office-status');
  const dotEl = statusEl.querySelector('.status-dot');
  statusEl.title = message;
  
  if (isConnected) {
    dotEl.className = 'status-dot online';
  } else {
    dotEl.className = 'status-dot';
  }
}

function simulateBrowserContext() {
  state.currentSlideInfo = {
    index: 1,
    title: 'Demo Presentation: AI in Enterprise Strategy',
    text: 'Key Drivers: Operational agility, automated workflows, intelligent decision making. Challenges: Legacy migration and data quality.',
    shapeCount: 2
  };
  updateSlideContextBadge();
}

// ============================================================================
// PowerPoint Integration (Office.js API)
// ============================================================================

/**
 * Reads text and structure from the currently active slide in PowerPoint.
 */
async function scanActiveSlide() {
  if (!state.isPowerPoint || typeof PowerPoint === 'undefined') {
    simulateBrowserContext();
    return;
  }

  const labelEl = document.getElementById('active-slide-label');
  labelEl.textContent = 'Scanning slide...';

  try {
    await PowerPoint.run(async (context) => {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load('items');
      await context.sync();

      if (!selectedSlides.items || selectedSlides.items.length === 0) {
        labelEl.textContent = 'No slide selected';
        return;
      }

      const slide = selectedSlides.items[0];
      const shapes = slide.shapes;
      shapes.load('items/textFrame/textRange/text, items/name, items/id');
      await context.sync();

      let extractedTexts = [];
      let slideTitle = '';

      for (const shape of shapes.items) {
        if (shape.textFrame && shape.textFrame.textRange && shape.textFrame.textRange.text) {
          const text = shape.textFrame.textRange.text.trim();
          if (text) {
            extractedTexts.push(text);
            if (!slideTitle) {
              slideTitle = text.split('\n')[0];
            }
          }
        }
      }

      state.currentSlideInfo = {
        index: 1,
        title: slideTitle || 'Current Slide',
        text: extractedTexts.join('\n\n'),
        shapeCount: shapes.items.length
      };

      updateSlideContextBadge();
    });
  } catch (error) {
    console.warn('Could not read slide via Office.js:', error);
    labelEl.textContent = 'Slide detected (ready)';
  }
}

function updateSlideContextBadge() {
  const labelEl = document.getElementById('active-slide-label');
  if (state.currentSlideInfo.text) {
    const preview = state.currentSlideInfo.title || 'Slide with text';
    const truncated = preview.length > 28 ? preview.substring(0, 26) + '...' : preview;
    labelEl.textContent = `Active: ${truncated} (${state.currentSlideInfo.shapeCount} items)`;
  } else {
    labelEl.textContent = 'Active slide: Empty or no text detected';
  }
}

/**
 * Replaces selected text in PowerPoint document.
 */
function replaceSelectionInDocument(text) {
  if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
    Office.context.document.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          showToast('Updated selected text in slide!');
        } else {
          showToast('Select a text box in PowerPoint first, then click Apply.');
        }
      }
    );
  } else {
    navigator.clipboard.writeText(text);
    showToast('Copied to clipboard (Preview mode)!');
  }
}

/**
 * Inserts a fully designed slide with Title, subtitle, and styled content cards or bullets.
 */
async function insertSlideIntoPowerPoint(slideData) {
  if (!state.isPowerPoint || typeof PowerPoint === 'undefined') {
    showToast('Simulated: Slide added in preview mode!');
    return;
  }

  try {
    await PowerPoint.run(async (context) => {
      // 1. Add a new blank slide
      const slides = context.presentation.slides;
      const newSlide = slides.add();

      // Standard 16:9 presentation dimensions (points): 960 x 540 or 720 x 405
      // PowerPoint default width is approx 960pt, height 540pt (or 720x405)
      const slideWidth = 720;
      const slideHeight = 405;

      // 2. Add Title Box
      const titleLeft = 40;
      const titleTop = 30;
      const titleWidth = slideWidth - 80;
      const titleHeight = 50;

      const titleShape = newSlide.shapes.addTextBox(slideData.title || 'Untitled Slide', {
        left: titleLeft,
        top: titleTop,
        width: titleWidth,
        height: titleHeight
      });
      titleShape.textFrame.textRange.font.size = 24;
      titleShape.textFrame.textRange.font.bold = true;
      titleShape.textFrame.textRange.font.color = '#0F172A';

      // 3. Add Subtitle or Category pill if provided
      if (slideData.subtitle) {
        const subShape = newSlide.shapes.addTextBox(slideData.subtitle, {
          left: titleLeft,
          top: titleTop + 45,
          width: titleWidth,
          height: 30
        });
        subShape.textFrame.textRange.font.size = 13;
        subShape.textFrame.textRange.font.color = '#2563EB';
      }

      // 4. Content cards or bullet points
      if (slideData.items && slideData.items.length > 0) {
        const count = Math.min(slideData.items.length, 3);
        const cardGap = 16;
        const totalGap = cardGap * (count - 1);
        const cardWidth = (titleWidth - totalGap) / count;
        const cardTop = titleTop + 85;
        const cardHeight = 220;

        for (let i = 0; i < count; i++) {
          const item = slideData.items[i];
          const cardLeft = titleLeft + i * (cardWidth + cardGap);

          // Add card shape
          const cardShape = newSlide.shapes.addTextBox(
            `${item.heading || `Key Point ${i+1}`}\n\n${item.description || item}`,
            {
              left: cardLeft,
              top: cardTop,
              width: cardWidth,
              height: cardHeight
            }
          );
          cardShape.textFrame.textRange.font.size = 12;
          cardShape.textFrame.textRange.font.color = '#334155';
        }
      } else if (slideData.bullets && slideData.bullets.length > 0) {
        // Bullet list layout
        const bulletText = slideData.bullets.map(b => `•  ${b}`).join('\n\n');
        const bulletsShape = newSlide.shapes.addTextBox(bulletText, {
          left: titleLeft,
          top: titleTop + 85,
          width: titleWidth,
          height: 240
        });
        bulletsShape.textFrame.textRange.font.size = 14;
        bulletsShape.textFrame.textRange.font.color = '#1E293B';
      }

      // 5. Takeaway footer if present
      if (slideData.takeaway) {
        const takeawayShape = newSlide.shapes.addTextBox(`💡 Key Takeaway: ${slideData.takeaway}`, {
          left: titleLeft,
          top: slideHeight - 45,
          width: titleWidth,
          height: 30
        });
        takeawayShape.textFrame.textRange.font.size = 11;
        takeawayShape.textFrame.textRange.font.italic = true;
        takeawayShape.textFrame.textRange.font.color = '#64748B';
      }

      await context.sync();
      showToast('✨ New slide inserted into your presentation!');
      scanActiveSlide();
    });
  } catch (error) {
    console.error('Error adding slide:', error);
    showToast('Failed to insert slide. Check PowerPoint permissions.');
  }
}

// ============================================================================
// Google Gemini API Client (Google AI Studio)
// ============================================================================

/**
 * Calls Google AI Studio Gemini API endpoint directly from the client.
 */
async function callGeminiApi({ prompt, systemInstruction, jsonOutput = false }) {
  if (!state.apiKey) {
    openSettingsModal();
    showToast('Please enter your Google AI Studio API key');
    throw new Error('API Key missing');
  }

  const model = state.model || 'gemini-3.8-flash';
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(state.apiKey)}`;

  // Assemble system instruction combining custom instructions + presentation specialist persona
  let fullSystemInstruction = `You are Gemini Copilot, an elite executive presentation designer and copywriter for Microsoft PowerPoint, working exactly like modern AI presentation assistants. You write punchy, clear, high-impact slide text, structured outlines, and natural speaker notes. Always keep presentation copy concise, structured, and visually engaging. Avoid dense text walls.`;

  if (state.customInstructions) {
    fullSystemInstruction += `\n\nUser Specific Directives: ${state.customInstructions}`;
  }

  if (systemInstruction) {
    fullSystemInstruction += `\n\nTask Directive: ${systemInstruction}`;
  }

  const requestBody = {
    contents: [
      {
        role: 'user',
        parts: [{ text: prompt }]
      }
    ],
    systemInstruction: {
      parts: [{ text: fullSystemInstruction }]
    },
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 2048,
      ...(jsonOutput ? { responseMimeType: 'application/json' } : {})
    }
  };

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      const msg = errorData.error?.message || `HTTP error ${response.status}: ${response.statusText}`;
      throw new Error(msg);
    }

    const data = await response.json();
    const candidate = data.candidates?.[0];
    const outputText = candidate?.content?.parts?.[0]?.text;

    if (!outputText) {
      throw new Error('No response text returned by Gemini.');
    }

    return outputText;
  } catch (error) {
    console.error('Gemini API Error:', error);
    throw error;
  }
}

// ============================================================================
// UI Actions & Feature Handlers
// ============================================================================

// 1. Chat Tab Handler
async function handleSendChatMessage() {
  const inputEl = document.getElementById('chat-input');
  const userText = inputEl.value.trim();
  if (!userText) return;

  // Add User Bubble
  appendChatMessage('user', userText);
  inputEl.value = '';
  inputEl.style.height = 'auto';

  // Include slide context if toggle checked
  const includeContext = document.getElementById('include-context-toggle').checked;
  let fullPrompt = userText;

  if (includeContext && state.currentSlideInfo.text) {
    fullPrompt = `[Context: Current active PowerPoint slide content]:\n"${state.currentSlideInfo.text}"\n\nUser Question/Instruction:\n${userText}`;
  }

  // Temporary Loading Bubble
  const loadingBubble = appendChatMessage('assistant', 'Thinking with Gemini...', true);

  try {
    const responseText = await callGeminiApi({ prompt: fullPrompt });
    updateAssistantBubble(loadingBubble, responseText);
  } catch (error) {
    updateAssistantBubble(loadingBubble, `⚠️ **Error calling Gemini API**: ${error.message}\n\nPlease check your API key in Settings.`);
  }
}

function appendChatMessage(role, text, isLoading = false) {
  const container = document.getElementById('chat-messages');
  const bubble = document.createElement('div');
  bubble.className = `message-bubble message-${role}`;

  const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

  if (role === 'user') {
    bubble.innerHTML = `
      <div class="bubble-content">${escapeHtml(text)}</div>
    `;
  } else {
    bubble.innerHTML = `
      <div class="bubble-header">
        <span class="bubble-sender">Gemini Copilot</span>
        <span class="bubble-time">${timeStr}</span>
      </div>
      <div class="bubble-content markdown-body">
        ${isLoading ? '<div class="pulse-indicator"></div> ' + escapeHtml(text) : renderMarkdown(text)}
      </div>
      ${!isLoading ? `
        <div class="bubble-actions">
          <button class="copy-small-btn btn-insert-chat" title="Insert response into PowerPoint">✨ Insert to Slide</button>
          <button class="copy-small-btn btn-copy-chat" title="Copy response to clipboard">Copy</button>
        </div>
      ` : ''}
    `;

    if (!isLoading) {
      attachBubbleActionListeners(bubble, text);
    }
  }

  container.appendChild(bubble);
  container.scrollTop = container.scrollHeight;
  return bubble;
}

function updateAssistantBubble(bubble, newText) {
  const contentEl = bubble.querySelector('.bubble-content');
  if (contentEl) {
    contentEl.innerHTML = renderMarkdown(newText);
  }

  let actionsEl = bubble.querySelector('.bubble-actions');
  if (!actionsEl) {
    actionsEl = document.createElement('div');
    actionsEl.className = 'bubble-actions';
    actionsEl.innerHTML = `
      <button class="copy-small-btn btn-insert-chat" title="Insert response into PowerPoint">✨ Insert to Slide</button>
      <button class="copy-small-btn btn-copy-chat" title="Copy response to clipboard">Copy</button>
    `;
    bubble.appendChild(actionsEl);
  }
  attachBubbleActionListeners(bubble, newText);

  const container = document.getElementById('chat-messages');
  container.scrollTop = container.scrollHeight;
}

function attachBubbleActionListeners(bubble, text) {
  const insertBtn = bubble.querySelector('.btn-insert-chat');
  const copyBtn = bubble.querySelector('.btn-copy-chat');

  if (insertBtn) {
    insertBtn.addEventListener('click', () => {
      // Create a slide from this chat content
      const lines = text.split('\n').filter(l => l.trim().length > 0);
      const title = lines[0].replace(/^#+\s*/, '').replace(/\*\*/g, '').substring(0, 60);
      const bullets = lines.slice(1, 6).map(l => l.replace(/^[-*•]\s*/, '').replace(/\*\*/g, ''));

      insertSlideIntoPowerPoint({
        title: title || 'Slide Summary',
        subtitle: 'Created with Gemini Copilot',
        bullets: bullets.length > 0 ? bullets : [text.substring(0, 150)]
      });
    });
  }

  if (copyBtn) {
    copyBtn.addEventListener('click', () => {
      navigator.clipboard.writeText(text);
      showToast('Copied to clipboard!');
    });
  }
}

// 2. Slide Creator Handler
async function handleGenerateSlide() {
  const topicInput = document.getElementById('slide-topic-input').value.trim();
  if (!topicInput) {
    showToast('Please enter a topic or description for the slide.');
    return;
  }

  const layout = document.getElementById('slide-layout-select').value;
  const includeNotes = document.getElementById('include-speaker-notes-checkbox').checked;
  const genBtn = document.getElementById('generate-slide-btn');

  genBtn.disabled = true;
  genBtn.innerHTML = `<span>Designing slide with Gemini...</span>`;

  const prompt = `
Create a professional, structured PowerPoint presentation slide for the following topic:
"${topicInput}"

Layout style requested: ${layout}.
Include speaker notes: ${includeNotes}.

You MUST return valid JSON adhering strictly to this schema:
{
  "title": "Punchy, executive slide title (max 8 words)",
  "subtitle": "Short context or category subtitle (max 6 words)",
  "items": [
    { "heading": "Point 1 Title", "description": "Concise high-impact explanation (1-2 sentences)" },
    { "heading": "Point 2 Title", "description": "Concise high-impact explanation (1-2 sentences)" },
    { "heading": "Point 3 Title", "description": "Concise high-impact explanation (1-2 sentences)" }
  ],
  "takeaway": "Single bottom line takeaway for the executive audience",
  "speakerNotes": "Natural talking points for the presenter"
}
`;

  try {
    const rawJson = await callGeminiApi({ prompt, jsonOutput: true });
    let slideData;
    try {
      slideData = JSON.parse(rawJson);
    } catch {
      // Fallback regex extract
      const jsonMatch = rawJson.match(/\{[\s\S]*\}/);
      if (jsonMatch) slideData = JSON.parse(jsonMatch[0]);
    }

    if (!slideData || !slideData.title) {
      throw new Error('Could not parse generated slide data.');
    }

    state.lastGeneratedSlide = slideData;
    renderSlidePreview(slideData);

    // Directly insert into PowerPoint
    await insertSlideIntoPowerPoint(slideData);
  } catch (error) {
    console.error(error);
    showToast(`Error generating slide: ${error.message}`);
  } finally {
    genBtn.disabled = false;
    genBtn.innerHTML = `
      <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon></svg>
      <span>Generate & Insert Slide</span>
    `;
  }
}

function renderSlidePreview(slideData) {
  const container = document.getElementById('slide-preview-container');
  const body = document.getElementById('slide-preview-body');

  let cardsHtml = '';
  if (slideData.items) {
    cardsHtml = slideData.items.map(item => `
      <div class="preview-card-item">
        <div class="preview-card-title">${escapeHtml(item.heading)}</div>
        <div class="preview-card-text">${escapeHtml(item.description)}</div>
      </div>
    `).join('');
  }

  body.innerHTML = `
    <div class="preview-slide-title">${escapeHtml(slideData.title)}</div>
    ${slideData.subtitle ? `<div style="font-size:12px; color:var(--accent-blue); margin-bottom:8px;">${escapeHtml(slideData.subtitle)}</div>` : ''}
    <div class="preview-cards-grid">
      ${cardsHtml}
    </div>
    ${slideData.takeaway ? `<div style="margin-top:10px; font-size:11px; font-weight:600; color:var(--text-secondary);">💡 Takeaway: ${escapeHtml(slideData.takeaway)}</div>` : ''}
    ${slideData.speakerNotes ? `<div class="preview-notes-box"><strong>Speaker Notes:</strong> ${escapeHtml(slideData.speakerNotes)}</div>` : ''}
  `;

  container.classList.remove('hidden');
}

// 3. Refine & Tools Handlers
async function handleRefineAction(actionType, customInstruction = '') {
  let targetText = '';

  // Attempt to read selected text or active slide text
  if (state.currentSlideInfo.text) {
    targetText = state.currentSlideInfo.text;
  } else {
    showToast('Please select a text box on the slide or scan the slide first.');
    return;
  }

  let instruction = '';
  switch (actionType) {
    case 'concise':
      instruction = 'Rewrite the following presentation text to be significantly more concise, punchy, and impactful for executive slides:';
      break;
    case 'executive':
      instruction = 'Elevate this text to an executive-ready, polished corporate presentation tone with active verbs:';
      break;
    case 'bullets':
      instruction = 'Convert the following slide text into 3-4 crisp, parallel bullet points with bold lead-ins:';
      break;
    case 'persuasive':
      instruction = 'Rewrite this text to be highly persuasive, compelling, and value-oriented:';
      break;
    case 'grammar':
      instruction = 'Fix all grammar, clarity, spelling, and phrasing issues while maintaining original presentation meaning:';
      break;
    case 'custom':
      instruction = `${customInstruction}:`;
      break;
  }

  const prompt = `${instruction}\n\n"${targetText}"\n\nReturn only the refined text ready to be inserted directly into the slide.`;

  const resultBox = document.getElementById('refine-result-box');
  const resultText = document.getElementById('refine-result-text');
  const replaceBtn = document.getElementById('refine-replace-btn');

  resultBox.classList.remove('hidden');
  resultText.textContent = 'Polishing with Gemini...';

  try {
    const refined = await callGeminiApi({ prompt });
    resultText.textContent = refined.trim();
    replaceBtn.classList.remove('hidden');

    // Attach click to replace
    replaceBtn.onclick = () => {
      replaceSelectionInDocument(refined.trim());
    };
  } catch (error) {
    resultText.textContent = `Error: ${error.message}`;
  }
}

async function handleGenerateNotes() {
  if (!state.currentSlideInfo.text) {
    showToast('No slide text detected. Write some text on your slide first!');
    return;
  }

  const prompt = `Based on this PowerPoint slide content:
"${state.currentSlideInfo.text}"

Generate natural, engaging talking points and speaker notes that the presenter can speak out loud. Format with 3-4 bullet points emphasizing tone, transitions, and key takeaways.`;

  const box = document.getElementById('notes-result-box');
  const textEl = document.getElementById('notes-result-text');
  box.classList.remove('hidden');
  textEl.textContent = 'Crafting speaker notes...';

  try {
    const notes = await callGeminiApi({ prompt });
    textEl.textContent = notes.trim();
  } catch (err) {
    textEl.textContent = `Error: ${err.message}`;
  }
}

async function handleCritiqueSlide() {
  if (!state.currentSlideInfo.text) {
    showToast('No slide text detected to analyze.');
    return;
  }

  const prompt = `Analyze this PowerPoint slide as an expert presentation design coach:
"${state.currentSlideInfo.text}"

Provide a structured critique covering:
1. **Clarity & Brevity** (Is it too wordy?)
2. **Audience Engagement** (Is the takeaway obvious in 3 seconds?)
3. **Actionable Suggestions** (3 direct improvements).`;

  const box = document.getElementById('critique-result-box');
  const textEl = document.getElementById('critique-result-text');
  box.classList.remove('hidden');
  textEl.innerHTML = 'Analyzing slide...';

  try {
    const critique = await callGeminiApi({ prompt });
    textEl.innerHTML = renderMarkdown(critique);
  } catch (err) {
    textEl.textContent = `Error: ${err.message}`;
  }
}

// ============================================================================
// Settings & Key Management
// ============================================================================

function maskApiKey(key) {
  if (!key) return '';
  if (key.length <= 8) return '••••••••';
  const start = key.substring(0, 6);
  const end = key.substring(key.length - 4);
  return `${start}••••••••••••••••${end}`;
}

function loadSavedSettings() {
  state.apiKey = localStorage.getItem('gemini_api_key') || '';
  state.model = localStorage.getItem('gemini_model') || 'gemini-3.8-flash';
  state.customInstructions = localStorage.getItem('gemini_instructions') || '';

  const keyInput = document.getElementById('api-key-input');
  const clearKeyBtn = document.getElementById('clear-key-btn');
  const securityNote = document.getElementById('key-security-note');
  const modelSelect = document.getElementById('model-select');
  const instructionsInput = document.getElementById('custom-instructions');
  const modelTag = document.getElementById('model-tag');

  if (state.apiKey) {
    keyInput.value = maskApiKey(state.apiKey);
    keyInput.disabled = true;
    keyInput.type = 'text'; // Safe because it's masked with bullets
    if (clearKeyBtn) clearKeyBtn.classList.remove('hidden');
    if (securityNote) {
      securityNote.innerHTML = '🔒 <strong>Key Saved & Protected:</strong> The raw key is masked and never exposed in the UI.';
      securityNote.style.color = 'var(--accent-green)';
    }
  } else {
    keyInput.value = '';
    keyInput.disabled = false;
    keyInput.type = 'password';
    if (clearKeyBtn) clearKeyBtn.classList.add('hidden');
    if (securityNote) {
      securityNote.innerHTML = 'Get your key from <a href="https://aistudio.google.com/app/apikey" target="_blank" rel="noopener">Google AI Studio</a>. Keys are stored locally on your machine and never sent anywhere except directly to Google\'s API.';
      securityNote.style.color = '';
    }
  }

  if (modelSelect) modelSelect.value = state.model;
  if (instructionsInput) instructionsInput.value = state.customInstructions;
  if (modelTag) modelTag.textContent = state.model;
}

function saveSettings() {
  const keyInput = document.getElementById('api-key-input');
  const modelSelect = document.getElementById('model-select');
  const instructionsInput = document.getElementById('custom-instructions');

  // If input was enabled and modified
  if (!keyInput.disabled) {
    const rawVal = keyInput.value.trim();
    if (rawVal && !rawVal.includes('••')) {
      state.apiKey = rawVal;
      localStorage.setItem('gemini_api_key', state.apiKey);
    } else if (!rawVal) {
      state.apiKey = '';
      localStorage.removeItem('gemini_api_key');
    }
  }

  state.model = modelSelect.value || 'gemini-3.8-flash';
  state.customInstructions = instructionsInput.value.trim();

  localStorage.setItem('gemini_model', state.model);
  localStorage.setItem('gemini_instructions', state.customInstructions);

  document.getElementById('model-tag').textContent = state.model;
  closeSettingsModal();
  loadSavedSettings();
  updateUI();
  showToast('Settings saved successfully!');
}

async function testApiKeyConnection() {
  const keyInput = document.getElementById('api-key-input');
  const statusMsg = document.getElementById('test-key-status');
  
  // Use state.apiKey if input is masked/disabled, otherwise use entered value
  const keyToTest = (keyInput.disabled || !keyInput.value.trim() || keyInput.value.includes('••')) 
    ? state.apiKey 
    : keyInput.value.trim();

  if (!keyToTest) {
    statusMsg.textContent = '❌ Please enter an API key first.';
    statusMsg.style.color = 'var(--accent-red)';
    return;
  }

  statusMsg.textContent = '⏳ Testing connection with gemini-3.8-flash...';
  statusMsg.style.color = 'var(--accent-blue)';

  try {
    const testUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=${encodeURIComponent(keyToTest)}`;
    const res = await fetch(testUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ role: 'user', parts: [{ text: 'Hello' }] }]
      })
    });

    if (res.ok) {
      statusMsg.textContent = '✅ Connected successfully to Google AI Studio (gemini-3.8-flash)!';
      statusMsg.style.color = 'var(--accent-green)';
    } else {
      const err = await res.json().catch(() => ({}));
      statusMsg.textContent = `❌ ${err.error?.message || 'Invalid API Key'}`;
      statusMsg.style.color = 'var(--accent-red)';
    }
  } catch (error) {
    statusMsg.textContent = `❌ Network Error: ${error.message}`;
    statusMsg.style.color = 'var(--accent-red)';
  }
}

function updateUI() {
  const banner = document.getElementById('api-key-banner');
  if (!state.apiKey) {
    banner.classList.remove('hidden');
  } else {
    banner.classList.add('hidden');
  }
}

function openSettingsModal() {
  document.getElementById('settings-modal').classList.remove('hidden');
}

function closeSettingsModal() {
  document.getElementById('settings-modal').classList.add('hidden');
  document.getElementById('test-key-status').textContent = '';
}

// ============================================================================
// Event Listeners Setup
// ============================================================================

function initDOMListeners() {
  // Navigation Tabs
  document.querySelectorAll('.nav-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.nav-tab').forEach(t => {
        t.classList.remove('active');
        t.setAttribute('aria-selected', 'false');
      });
      document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));

      tab.classList.add('active');
      tab.setAttribute('aria-selected', 'true');
      const targetPane = document.getElementById(tab.dataset.tab);
      if (targetPane) targetPane.classList.add('active');
    });
  });

  // Settings Modal
  document.getElementById('open-settings-btn').addEventListener('click', openSettingsModal);
  document.getElementById('close-settings-btn').addEventListener('click', closeSettingsModal);
  document.getElementById('banner-setup-btn').addEventListener('click', openSettingsModal);
  document.getElementById('save-settings-btn').addEventListener('click', saveSettings);
  document.getElementById('test-key-btn').addEventListener('click', testApiKeyConnection);

  // Change/Clear API Key action
  const clearKeyBtn = document.getElementById('clear-key-btn');
  const keyInput = document.getElementById('api-key-input');
  const securityNote = document.getElementById('key-security-note');

  if (clearKeyBtn) {
    clearKeyBtn.addEventListener('click', () => {
      keyInput.disabled = false;
      keyInput.value = '';
      keyInput.type = 'password';
      keyInput.focus();
      clearKeyBtn.classList.add('hidden');
      if (securityNote) {
        securityNote.innerHTML = 'Enter a new Google AI Studio API key and click <strong>Save Settings</strong>.';
        securityNote.style.color = 'var(--text-secondary)';
      }
    });
  }

  // Scan Slide Button
  document.getElementById('refresh-slide-btn').addEventListener('click', () => {
    scanActiveSlide();
    showToast('Slide context refreshed');
  });

  // Chat input
  const chatInput = document.getElementById('chat-input');
  const sendBtn = document.getElementById('chat-send-btn');
  sendBtn.addEventListener('click', handleSendChatMessage);

  chatInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSendChatMessage();
    }
  });

  // Auto-resize chat textarea
  chatInput.addEventListener('input', () => {
    chatInput.style.height = 'auto';
    chatInput.style.height = `${Math.min(chatInput.scrollHeight, 100)}px`;
  });

  // Clear chat
  document.getElementById('clear-chat-btn').addEventListener('click', () => {
    const container = document.getElementById('chat-messages');
    container.innerHTML = `
      <div class="message-bubble message-assistant">
        <div class="bubble-header"><span class="bubble-sender">Gemini Copilot</span></div>
        <div class="bubble-content"><p>Conversation cleared. How can I help with your slides?</p></div>
      </div>
    `;
  });

  // Chat Suggestion Chips
  document.querySelectorAll('.chip-btn').forEach(chip => {
    chip.addEventListener('click', () => {
      const prompt = chip.dataset.prompt;
      document.getElementById('chat-input').value = prompt;
      handleSendChatMessage();
    });
  });

  // Create Slide Tab
  document.getElementById('generate-slide-btn').addEventListener('click', handleGenerateSlide);
  document.getElementById('insert-preview-btn').addEventListener('click', () => {
    if (state.lastGeneratedSlide) {
      insertSlideIntoPowerPoint(state.lastGeneratedSlide);
    }
  });

  // Refine Tab Presets
  document.querySelectorAll('.preset-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const action = btn.dataset.action;
      handleRefineAction(action);
    });
  });

  document.getElementById('refine-execute-btn').addEventListener('click', () => {
    const custom = document.getElementById('refine-custom-input').value.trim();
    if (custom) {
      handleRefineAction('custom', custom);
    } else {
      handleRefineAction('concise');
    }
  });

  document.getElementById('copy-refine-btn').addEventListener('click', () => {
    const text = document.getElementById('refine-result-text').textContent;
    navigator.clipboard.writeText(text);
    showToast('Copied refined text!');
  });

  // Speaker Notes & Critique
  document.getElementById('gen-notes-btn').addEventListener('click', handleGenerateNotes);
  document.getElementById('critique-slide-btn').addEventListener('click', handleCritiqueSlide);
}

// ============================================================================
// Utilities & Toast
// ============================================================================

function showToast(message, duration = 3000) {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.classList.remove('hidden');

  setTimeout(() => {
    toast.classList.add('hidden');
  }, duration);
}

function escapeHtml(text) {
  const map = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' };
  return String(text).replace(/[&<>"']/g, m => map[m]);
}

function renderMarkdown(md) {
  if (!md) return '';
  let html = escapeHtml(md);

  // Bold **text**
  html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
  // Italic *text*
  html = html.replace(/\*(.*?)\*/g, '<em>$1</em>');
  // Inline code `code`
  html = html.replace(/`([^`]+)`/g, '<code>$1</code>');
  // Bullet lists (- or * at start of line)
  html = html.replace(/^\s*[-*•]\s+(.*)$/gm, '<li>$1</li>');
  html = html.replace(/(<li>.*<\/li>)/s, '<ul>$1</ul>');
  // Linebreaks
  html = html.replace(/\n\n/g, '<p></p>');
  html = html.replace(/\n/g, '<br/>');

  return html;
}
