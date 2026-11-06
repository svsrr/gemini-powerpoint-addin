/**
 * Gemini Copilot for PowerPoint
 * Powered by Google Gemini API from Google AI Studio
 */

// Application State
const getInitialModel = () => {
  const m = localStorage.getItem('gemini_model');
  if (!m || m.includes('2.5') || m.includes('2.0')) {
    localStorage.setItem('gemini_model', 'gemini-3.8-flash');
    return 'gemini-3.8-flash';
  }
  return m;
};

// Theme Color Palettes (Applied to generated slide shapes)
const THEMES = {
  'modern-blue': {
    name: 'Modern Blue & Indigo',
    titleColor: '#0F172A',
    subtitleColor: '#2563EB',
    cardBg: '#F8FAFC',
    cardBorder: '#BFDBFE',
    headingColor: '#1D4ED8',
    textColor: '#334155',
    takeawayColor: '#1E40AF',
    accent: '#2563EB'
  },
  'emerald-green': {
    name: 'Emerald Growth',
    titleColor: '#064E3B',
    subtitleColor: '#059669',
    cardBg: '#F0FDF4',
    cardBorder: '#A7F3D0',
    headingColor: '#047857',
    textColor: '#1F2937',
    takeawayColor: '#065F46',
    accent: '#10B981'
  },
  'dark-slate': {
    name: 'Executive Dark Slate',
    titleColor: '#0F172A',
    subtitleColor: '#64748B',
    cardBg: '#F1F5F9',
    cardBorder: '#CBD5E1',
    headingColor: '#0F172A',
    textColor: '#334155',
    takeawayColor: '#334155',
    accent: '#0284C7'
  },
  'coral-accent': {
    name: 'Warm Coral',
    titleColor: '#7F1D1D',
    subtitleColor: '#DC2626',
    cardBg: '#FFF5F5',
    cardBorder: '#FECACA',
    headingColor: '#B91C1C',
    textColor: '#374151',
    takeawayColor: '#991B1B',
    accent: '#EF4444'
  }
};

const state = {
  apiKey: localStorage.getItem('gemini_api_key') || '',
  model: getInitialModel(),
  customInstructions: localStorage.getItem('gemini_instructions') || '',
  isOfficeReady: false,
  isPowerPoint: false,
  currentSlideInfo: {
    index: 1,
    title: '',
    text: '',
    shapeCount: 0
  },
  chatHistory: [], // Full multi-turn conversation memory: [{ role: 'user'|'model', parts: [{ text }] }]
  lastGeneratedSlide: null,
  lastGeneratedNotes: ''
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
  if (!statusEl) return;
  const dotEl = statusEl.querySelector('.status-dot');
  statusEl.title = message;
  
  if (dotEl) {
    dotEl.className = isConnected ? 'status-dot online' : 'status-dot';
  }
}

function simulateBrowserContext() {
  state.currentSlideInfo = {
    index: 1,
    title: 'Executive AI Strategy Overview',
    text: 'Key Priorities: Accelerate enterprise workflows, maintain zero-loss data governance, and scale agentic intelligence across global business units.',
    shapeCount: 3
  };
  updateSlideContextBadge();
}

// ============================================================================
// Robust Slide Scanning & Office.js Helper Methods
// ============================================================================

/**
 * Reads text and structure safely from the currently active slide.
 * Avoids throwing on non-text shapes (images, lines, groups).
 */
async function scanActiveSlide() {
  if (!state.isPowerPoint || typeof PowerPoint === 'undefined') {
    simulateBrowserContext();
    return;
  }

  const labelEl = document.getElementById('active-slide-label');
  if (labelEl) labelEl.textContent = 'Scanning slide...';

  try {
    await PowerPoint.run(async (context) => {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load('items');
      await context.sync();

      if (!selectedSlides.items || selectedSlides.items.length === 0) {
        if (labelEl) labelEl.textContent = 'No slide selected';
        return;
      }

      const slide = selectedSlides.items[0];
      const shapes = slide.shapes;
      shapes.load('items/type, items/name, items/id, items/hasTextFrame');
      await context.sync();

      let extractedTexts = [];
      let slideTitle = '';

      for (const shape of shapes.items) {
        try {
          if (shape.hasTextFrame) {
            shape.textFrame.load('textRange/text');
            await context.sync();
            const text = shape.textFrame.textRange.text?.trim();
            if (text) {
              extractedTexts.push(text);
              if (!slideTitle) {
                slideTitle = text.split('\n')[0].replace(/^[#*•\s]+/, '');
              }
            }
          }
        } catch {
          // Gracefully skip shapes that do not support textRange
        }
      }

      state.currentSlideInfo = {
        index: 1,
        title: slideTitle || 'Active Slide',
        text: extractedTexts.join('\n\n'),
        shapeCount: shapes.items.length
      };

      updateSlideContextBadge();
    });
  } catch (error) {
    console.warn('Could not read slide via Office.js:', error);
    if (labelEl) labelEl.textContent = 'Slide detected (ready)';
  }
}

function updateSlideContextBadge() {
  const labelEl = document.getElementById('active-slide-label');
  if (!labelEl) return;
  if (state.currentSlideInfo.text) {
    const preview = state.currentSlideInfo.title || 'Slide with text';
    const truncated = preview.length > 28 ? preview.substring(0, 26) + '...' : preview;
    labelEl.textContent = `Active: ${truncated} (${state.currentSlideInfo.shapeCount} shapes)`;
  } else {
    labelEl.textContent = 'Active slide: Empty or no text detected';
  }
}

/**
 * Gets exact highlighted/selected text from PowerPoint document.
 */
function getSelectedTextFromDocument() {
  return new Promise((resolve) => {
    if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value && String(result.value).trim()) {
          resolve(String(result.value).trim());
        } else {
          resolve(null);
        }
      });
    } else {
      resolve(null);
    }
  });
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

// ============================================================================
// Full 16:9 Widescreen (960 x 540 pt) Slide Insertion & Layout Engine
// ============================================================================

/**
 * Inserts a newly designed 16:9 widescreen slide into PowerPoint.
 * Cleans out default empty placeholders and applies the chosen visual theme.
 */
async function insertSlideIntoPowerPoint(slideData) {
  if (!state.isPowerPoint || typeof PowerPoint === 'undefined') {
    showToast('Simulated: Slide added in browser preview mode!');
    return;
  }

  // Widescreen 16:9 standard dimensions in points
  const SLIDE_WIDTH = 960;
  const SLIDE_HEIGHT = 540;
  const MARGIN_LEFT = 60;
  const CONTENT_WIDTH = SLIDE_WIDTH - (MARGIN_LEFT * 2); // 840 pt

  const themeKey = slideData.theme || document.getElementById('slide-theme-select')?.value || 'modern-blue';
  const theme = THEMES[themeKey] || THEMES['modern-blue'];

  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      const newSlide = slides.add();

      // Clean out default template ghost placeholders ("Click to add title")
      newSlide.shapes.load('items');
      await context.sync();
      for (const shape of newSlide.shapes.items) {
        shape.delete();
      }
      await context.sync();

      // 1. Title Box
      const titleShape = newSlide.shapes.addTextBox(slideData.title || 'Executive Presentation', {
        left: MARGIN_LEFT,
        top: 40,
        width: CONTENT_WIDTH,
        height: 55
      });
      titleShape.textFrame.textRange.font.size = 28;
      titleShape.textFrame.textRange.font.bold = true;
      titleShape.textFrame.textRange.font.color = theme.titleColor;

      // 2. Subtitle / Category
      if (slideData.subtitle) {
        const subShape = newSlide.shapes.addTextBox(slideData.subtitle.toUpperCase(), {
          left: MARGIN_LEFT,
          top: 98,
          width: CONTENT_WIDTH,
          height: 25
        });
        subShape.textFrame.textRange.font.size = 12;
        subShape.textFrame.textRange.font.bold = true;
        subShape.textFrame.textRange.font.color = theme.subtitleColor;
      }

      // 3. Layout Rendering
      const layout = slideData.layout || 'cards';
      const contentTop = slideData.subtitle ? 135 : 115;
      const contentHeight = 310;

      if (layout === 'cards' || !layout) {
        // 3 Column Cards
        const items = slideData.items || [];
        const count = Math.max(1, Math.min(items.length, 3));
        const gap = 20;
        const cardWidth = (CONTENT_WIDTH - (gap * (count - 1))) / count;

        for (let i = 0; i < count; i++) {
          const item = items[i];
          const cardLeft = MARGIN_LEFT + i * (cardWidth + gap);

          const cardText = `${item.heading || `Point ${i+1}`}\n\n${item.description || item}`;
          const cardShape = newSlide.shapes.addTextBox(cardText, {
            left: cardLeft,
            top: contentTop,
            width: cardWidth,
            height: contentHeight
          });

          cardShape.textFrame.textRange.font.size = 13;
          cardShape.textFrame.textRange.font.color = theme.textColor;
        }
      } else if (layout === 'problem-solution') {
        // 2 Side-by-Side Comparison Cards
        const cardWidth = (CONTENT_WIDTH - 24) / 2;
        const problemText = `THE CHALLENGE\n\n${slideData.problem || (slideData.items?.[0]?.description || 'Challenge description')}`;
        const solutionText = `THE SOLUTION\n\n${slideData.solution || (slideData.items?.[1]?.description || 'Solution strategy')}`;

        const problemShape = newSlide.shapes.addTextBox(problemText, {
          left: MARGIN_LEFT,
          top: contentTop,
          width: cardWidth,
          height: contentHeight
        });
        problemShape.textFrame.textRange.font.size = 14;
        problemShape.textFrame.textRange.font.color = '#7F1D1D';

        const solutionShape = newSlide.shapes.addTextBox(solutionText, {
          left: MARGIN_LEFT + cardWidth + 24,
          top: contentTop,
          width: cardWidth,
          height: contentHeight
        });
        solutionShape.textFrame.textRange.font.size = 14;
        solutionShape.textFrame.textRange.font.color = '#064E3B';

      } else {
        // Bullet List / Timeline / Other
        const bullets = slideData.bullets || (slideData.items ? slideData.items.map(it => `${it.heading}: ${it.description}`) : []);
        const bulletText = bullets.map(b => `•  ${b}`).join('\n\n');

        const bulletsShape = newSlide.shapes.addTextBox(bulletText, {
          left: MARGIN_LEFT,
          top: contentTop,
          width: CONTENT_WIDTH,
          height: contentHeight
        });
        bulletsShape.textFrame.textRange.font.size = 15;
        bulletsShape.textFrame.textRange.font.color = theme.textColor;
      }

      // 4. Bottom Executive Takeaway
      if (slideData.takeaway) {
        const takeawayShape = newSlide.shapes.addTextBox(`💡 Bottom Line: ${slideData.takeaway}`, {
          left: MARGIN_LEFT,
          top: SLIDE_HEIGHT - 65,
          width: CONTENT_WIDTH,
          height: 35
        });
        takeawayShape.textFrame.textRange.font.size = 12;
        takeawayShape.textFrame.textRange.font.bold = true;
        takeawayShape.textFrame.textRange.font.color = theme.takeawayColor;
      }

      await context.sync();
      showToast('✨ Widescreen slide inserted into presentation!');
      scanActiveSlide();
    });
  } catch (error) {
    console.error('Error inserting slide:', error);
    showToast(`Error inserting slide: ${error.message}`);
  }
}

// ============================================================================
// Secure Google Gemini API Client (Using Header Auth & Multi-Turn Memory)
// ============================================================================

/**
 * Calls Google AI Studio Gemini API endpoint directly from client.
 * Uses 'x-goog-api-key' header for security (no keys in URL query string).
 */
async function callGeminiApi({ prompt, contents, systemInstruction, jsonSchema = null, temperature = 0.7 }) {
  if (!state.apiKey) {
    openSettingsModal();
    showToast('Please enter your Google AI Studio API key');
    throw new Error('API Key missing');
  }

  const model = state.model || 'gemini-3.8-flash';
  // Send key via header, NOT query parameter
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;

  let fullSystemInstruction = `You are Gemini Copilot, an elite executive presentation designer and copywriter for Microsoft PowerPoint, working exactly like modern AI presentation assistants. You write punchy, clear, high-impact slide text, structured outlines, and natural speaker notes. Always keep presentation copy concise, structured, and visually engaging. Avoid dense text walls.`;

  if (state.customInstructions) {
    fullSystemInstruction += `\n\nUser Directives: ${state.customInstructions}`;
  }
  if (systemInstruction) {
    fullSystemInstruction += `\n\nTask Directive: ${systemInstruction}`;
  }

  // Multi-turn contents array or single turn
  const requestContents = contents || [
    {
      role: 'user',
      parts: [{ text: prompt }]
    }
  ];

  const requestBody = {
    contents: requestContents,
    systemInstruction: {
      parts: [{ text: fullSystemInstruction }]
    },
    generationConfig: {
      temperature: temperature,
      maxOutputTokens: 4096,
      ...(jsonSchema ? {
        responseMimeType: 'application/json',
        responseSchema: jsonSchema
      } : {})
    }
  };

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-goog-api-key': state.apiKey
      },
      body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      const msg = errorData.error?.message || `HTTP ${response.status}: ${response.statusText}`;

      // Auto-fallback if the previously selected model was deprecated
      if ((msg.includes('no longer available') || msg.includes('not found')) && model !== 'gemini-3.8-flash') {
        console.warn(`Model ${model} deprecated, switching automatically to gemini-3.8-flash...`);
        state.model = 'gemini-3.8-flash';
        localStorage.setItem('gemini_model', 'gemini-3.8-flash');
        const tag = document.getElementById('model-tag');
        if (tag) tag.textContent = 'gemini-3.8-flash';

        const retryRes = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-goog-api-key': state.apiKey
          },
          body: JSON.stringify(requestBody)
        });

        if (retryRes.ok) {
          const retryData = await retryRes.json();
          const candidate = retryData.candidates?.[0];
          const textParts = candidate?.content?.parts?.map(p => p.text || '').join('');
          if (textParts) return textParts;
        }
      }

      throw new Error(msg);
    }

    const data = await response.json();
    const candidate = data.candidates?.[0];

    if (candidate?.finishReason && candidate.finishReason !== 'STOP') {
      console.warn('Gemini response finishReason:', candidate.finishReason);
      if (candidate.finishReason === 'SAFETY') {
        throw new Error('Response blocked by Gemini safety filters.');
      }
    }

    const outputText = candidate?.content?.parts?.map(p => p.text || '').join('');
    if (!outputText) {
      throw new Error('No content returned by Gemini.');
    }

    return outputText;
  } catch (error) {
    console.error('Gemini API Error:', error);
    throw error;
  }
}

// ============================================================================
// Multi-Turn Copilot Chat (With Memory & Follow-up Context)
// ============================================================================

async function handleSendChatMessage() {
  const inputEl = document.getElementById('chat-input');
  const userText = inputEl.value.trim();
  if (!userText) return;

  // Add User Bubble to UI
  appendChatMessage('user', userText);
  inputEl.value = '';
  inputEl.style.height = 'auto';

  // Inject current slide context if toggle checked
  const includeContext = document.getElementById('include-context-toggle')?.checked;
  let turnText = userText;

  if (includeContext && state.currentSlideInfo.text && state.chatHistory.length === 0) {
    turnText = `[Current Active Slide Context: "${state.currentSlideInfo.text}"]\n\n${userText}`;
  }

  // Append user message to multi-turn conversation memory
  state.chatHistory.push({
    role: 'user',
    parts: [{ text: turnText }]
  });

  const loadingBubble = appendChatMessage('assistant', 'Thinking with Gemini 3.8 Flash...', true);

  try {
    const responseText = await callGeminiApi({
      contents: state.chatHistory,
      temperature: 0.7
    });

    // Save assistant reply to memory for seamless follow-ups ("make it shorter", etc.)
    state.chatHistory.push({
      role: 'model',
      parts: [{ text: responseText }]
    });

    updateAssistantBubble(loadingBubble, responseText);
  } catch (error) {
    // Pop failed user turn from memory
    state.chatHistory.pop();
    updateAssistantBubble(loadingBubble, `⚠️ **Error calling Gemini API**: ${error.message}\n\nPlease check your API key in Settings.`);
  }
}

function appendChatMessage(role, text, isLoading = false) {
  const container = document.getElementById('chat-messages');
  if (!container) return null;
  const bubble = document.createElement('div');
  bubble.className = `message-bubble message-${role}`;

  const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

  if (role === 'user') {
    bubble.innerHTML = `<div class="bubble-content">${escapeHtml(text)}</div>`;
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
          <button class="copy-small-btn btn-insert-chat" title="Create slide from this response">✨ Insert to Slide</button>
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
      <button class="copy-small-btn btn-insert-chat" title="Create slide from this response">✨ Insert to Slide</button>
      <button class="copy-small-btn btn-copy-chat" title="Copy response to clipboard">Copy</button>
    `;
    bubble.appendChild(actionsEl);
  }
  attachBubbleActionListeners(bubble, newText);

  const container = document.getElementById('chat-messages');
  if (container) container.scrollTop = container.scrollHeight;
}

function attachBubbleActionListeners(bubble, text) {
  const insertBtn = bubble.querySelector('.btn-insert-chat');
  const copyBtn = bubble.querySelector('.btn-copy-chat');

  if (insertBtn) {
    insertBtn.addEventListener('click', () => {
      // Parse markdown structured blocks or lines
      const cleanLines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
      const titleLine = cleanLines.find(l => l.startsWith('#') || l.startsWith('**')) || cleanLines[0] || 'Presentation Slide';
      const cleanTitle = titleLine.replace(/^[#*•\s]+/, '').replace(/\*\*/g, '').substring(0, 60);

      const bullets = cleanLines
        .filter(l => l !== titleLine && (l.startsWith('-') || l.startsWith('*') || l.startsWith('•') || /^\d+\./.test(l)))
        .map(l => l.replace(/^[-*•\d.]+\s*/, '').replace(/\*\*/g, ''));

      insertSlideIntoPowerPoint({
        title: cleanTitle,
        subtitle: 'Created with Gemini Copilot',
        bullets: bullets.length > 0 ? bullets : cleanLines.slice(1, 5),
        layout: 'bullet-list'
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

// ============================================================================
// Slide Creator: Preview-First Generation (No Immediate Insertion)
// ============================================================================

async function handleGenerateSlide() {
  const topicInput = document.getElementById('slide-topic-input').value.trim();
  if (!topicInput) {
    showToast('Please enter a topic or description for the slide.');
    return;
  }

  const layout = document.getElementById('slide-layout-select')?.value || 'cards';
  const theme = document.getElementById('slide-theme-select')?.value || 'modern-blue';
  const includeNotes = document.getElementById('include-speaker-notes-checkbox')?.checked ?? true;
  const genBtn = document.getElementById('generate-slide-btn');

  genBtn.disabled = true;
  genBtn.innerHTML = `<span>Designing slide with Gemini...</span>`;

  // Structured Output Schema
  const slideSchema = {
    type: "OBJECT",
    properties: {
      title: { type: "STRING", description: "Punchy, executive slide title (max 8 words)" },
      subtitle: { type: "STRING", description: "Category or subhead (max 5 words)" },
      layout: { type: "STRING", description: layout },
      items: {
        type: "ARRAY",
        items: {
          type: "OBJECT",
          properties: {
            heading: { type: "STRING" },
            description: { type: "STRING" }
          },
          required: ["heading", "description"]
        }
      },
      bullets: {
        type: "ARRAY",
        items: { type: "STRING" }
      },
      problem: { type: "STRING" },
      solution: { type: "STRING" },
      takeaway: { type: "STRING", description: "Bottom line takeaway" },
      speakerNotes: { type: "STRING", description: "Talking points for presenter" }
    },
    required: ["title"]
  };

  const prompt = `
Create a high-impact PowerPoint presentation slide based on this topic:
"${topicInput}"

Layout requested: "${layout}".
Theme requested: "${theme}".
Include speaker notes: ${includeNotes}.

For cards or timeline layouts: Provide 3 items with clear heading and 1-2 sentence description.
For bullet-list layout: Provide 4-5 crisp, executive bullets with bold lead-ins.
For problem-solution layout: Provide problem and solution descriptions.
Provide a clear, decisive bottom-line takeaway.
`;

  try {
    const rawJson = await callGeminiApi({
      prompt,
      jsonSchema: slideSchema,
      temperature: 0.2 // Low temperature for clean structured outputs
    });

    let slideData;
    try {
      slideData = JSON.parse(rawJson);
    } catch {
      const match = rawJson.match(/\{[\s\S]*\}/);
      if (match) slideData = JSON.parse(match[0]);
    }

    if (!slideData || !slideData.title) {
      throw new Error('Failed to parse generated slide data.');
    }

    slideData.theme = theme;
    slideData.layout = layout;
    state.lastGeneratedSlide = slideData;

    // Render Preview First (DO NOT insert into PowerPoint yet!)
    renderSlidePreview(slideData);
    showToast('Slide generated! Review preview below, then click Insert.');
  } catch (error) {
    console.error(error);
    showToast(`Error generating slide: ${error.message}`);
  } finally {
    genBtn.disabled = false;
    genBtn.innerHTML = `
      <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon></svg>
      <span>Generate Slide Preview</span>
    `;
  }
}

function renderSlidePreview(slideData) {
  const container = document.getElementById('slide-preview-container');
  const body = document.getElementById('slide-preview-body');
  if (!container || !body) return;

  let contentHtml = '';
  if (slideData.items && slideData.items.length > 0) {
    contentHtml = `
      <div class="preview-cards-grid">
        ${slideData.items.map(item => `
          <div class="preview-card-item">
            <div class="preview-card-title">${escapeHtml(item.heading)}</div>
            <div class="preview-card-text">${escapeHtml(item.description)}</div>
          </div>
        `).join('')}
      </div>
    `;
  } else if (slideData.bullets && slideData.bullets.length > 0) {
    contentHtml = `
      <ul style="padding-left:18px; margin-top:8px; font-size:12px;">
        ${slideData.bullets.map(b => `<li style="margin-bottom:4px;">${escapeHtml(b)}</li>`).join('')}
      </ul>
    `;
  }

  body.innerHTML = `
    <div class="preview-slide-title">${escapeHtml(slideData.title)}</div>
    ${slideData.subtitle ? `<div style="font-size:11px; font-weight:600; color:var(--accent-blue); margin-bottom:8px;">${escapeHtml(slideData.subtitle)}</div>` : ''}
    ${contentHtml}
    ${slideData.takeaway ? `<div style="margin-top:10px; font-size:11px; font-weight:600; color:var(--text-secondary);">💡 Takeaway: ${escapeHtml(slideData.takeaway)}</div>` : ''}
    ${slideData.speakerNotes ? `<div class="preview-notes-box"><strong>Speaker Notes:</strong> ${escapeHtml(slideData.speakerNotes)}</div>` : ''}
  `;

  container.classList.remove('hidden');
  container.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// ============================================================================
// Refine Tools: Real Selected Text Inspection & Notes Handlers
// ============================================================================

async function handleRefineAction(actionType, customInstruction = '') {
  // 1. Try reading the user's actual highlighted/selected text in PowerPoint!
  const selectedText = await getSelectedTextFromDocument();
  let targetText = selectedText;

  if (!targetText) {
    // Fallback to active slide text if no text is highlighted
    if (state.currentSlideInfo.text) {
      targetText = state.currentSlideInfo.text;
      showToast('Refining full slide text (highlight a specific text box to refine only that selection).');
    } else {
      showToast('Please select text inside a slide text box or scan the slide first.');
      return;
    }
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

  const prompt = `${instruction}\n\n"${targetText}"\n\nReturn ONLY the refined text ready to be inserted directly into the slide.`;

  const resultBox = document.getElementById('refine-result-box');
  const resultText = document.getElementById('refine-result-text');
  const replaceBtn = document.getElementById('refine-replace-btn');

  resultBox.classList.remove('hidden');
  resultText.textContent = 'Polishing with Gemini...';

  try {
    const refined = await callGeminiApi({ prompt, temperature: 0.3 });
    resultText.textContent = refined.trim();
    replaceBtn.classList.remove('hidden');

    // Replace ONLY the selected text in the document
    replaceBtn.onclick = () => {
      replaceSelectionInDocument(refined.trim());
    };
  } catch (error) {
    resultText.textContent = `Error: ${error.message}`;
  }
}

async function handleGenerateNotes() {
  let targetText = await getSelectedTextFromDocument();
  if (!targetText) {
    targetText = state.currentSlideInfo.text;
  }

  if (!targetText) {
    showToast('No slide text detected. Add text to your slide or scan the slide first!');
    return;
  }

  const prompt = `Based on this PowerPoint slide content:
"${targetText}"

Generate natural, engaging talking points and speaker notes that the presenter can speak out loud. Format with 3-4 bullet points emphasizing tone, transitions, and key takeaways.`;

  const box = document.getElementById('notes-result-box');
  const textEl = document.getElementById('notes-result-text');
  box.classList.remove('hidden');
  textEl.textContent = 'Crafting speaker notes...';

  try {
    const notes = await callGeminiApi({ prompt, temperature: 0.5 });
    const cleanNotes = notes.trim();
    state.lastGeneratedNotes = cleanNotes;
    textEl.textContent = cleanNotes;
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
  state.model = getInitialModel();
  state.customInstructions = localStorage.getItem('gemini_instructions') || '';

  const keyInput = document.getElementById('api-key-input');
  const clearKeyBtn = document.getElementById('clear-key-btn');
  const securityNote = document.getElementById('key-security-note');
  const modelSelect = document.getElementById('model-select');
  const instructionsInput = document.getElementById('custom-instructions');
  const modelTag = document.getElementById('model-tag');

  if (state.apiKey) {
    if (keyInput) {
      keyInput.value = maskApiKey(state.apiKey);
      keyInput.disabled = true;
      keyInput.type = 'text';
    }
    if (clearKeyBtn) clearKeyBtn.classList.remove('hidden');
    if (securityNote) {
      securityNote.innerHTML = '🔒 <strong>Key Saved & Protected:</strong> The raw key is masked and never exposed in the UI.';
      securityNote.style.color = 'var(--accent-green)';
    }
  } else {
    if (keyInput) {
      keyInput.value = '';
      keyInput.disabled = false;
      keyInput.type = 'password';
    }
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

  state.model = modelSelect?.value || 'gemini-3.8-flash';
  state.customInstructions = instructionsInput?.value.trim() || '';

  localStorage.setItem('gemini_model', state.model);
  localStorage.setItem('gemini_instructions', state.customInstructions);

  const modelTag = document.getElementById('model-tag');
  if (modelTag) modelTag.textContent = state.model;

  closeSettingsModal();
  loadSavedSettings();
  updateUI();
  showToast('Settings saved successfully!');
}

async function testApiKeyConnection() {
  const keyInput = document.getElementById('api-key-input');
  const statusMsg = document.getElementById('test-key-status');
  
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
    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-goog-api-key': keyToTest
      },
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
  if (!banner) return;
  if (!state.apiKey) {
    banner.classList.remove('hidden');
  } else {
    banner.classList.add('hidden');
  }
}

function openSettingsModal() {
  document.getElementById('settings-modal')?.classList.remove('hidden');
}

function closeSettingsModal() {
  document.getElementById('settings-modal')?.classList.add('hidden');
  const statusEl = document.getElementById('test-key-status');
  if (statusEl) statusEl.textContent = '';
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
  document.getElementById('open-settings-btn')?.addEventListener('click', openSettingsModal);
  document.getElementById('close-settings-btn')?.addEventListener('click', closeSettingsModal);
  document.getElementById('banner-setup-btn')?.addEventListener('click', openSettingsModal);
  document.getElementById('save-settings-btn')?.addEventListener('click', saveSettings);
  document.getElementById('test-key-btn')?.addEventListener('click', testApiKeyConnection);

  // Clear/Change API Key button
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
  document.getElementById('refresh-slide-btn')?.addEventListener('click', () => {
    scanActiveSlide();
    showToast('Slide context refreshed');
  });

  // Chat input
  const chatInput = document.getElementById('chat-input');
  const sendBtn = document.getElementById('chat-send-btn');
  if (sendBtn) sendBtn.addEventListener('click', handleSendChatMessage);

  if (chatInput) {
    chatInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        handleSendChatMessage();
      }
    });

    chatInput.addEventListener('input', () => {
      chatInput.style.height = 'auto';
      chatInput.style.height = `${Math.min(chatInput.scrollHeight, 100)}px`;
    });
  }

  // Clear chat (resets UI and resets multi-turn chat memory)
  document.getElementById('clear-chat-btn')?.addEventListener('click', () => {
    state.chatHistory = [];
    const container = document.getElementById('chat-messages');
    if (container) {
      container.innerHTML = `
        <div class="message-bubble message-assistant">
          <div class="bubble-header"><span class="bubble-sender">Gemini Copilot</span></div>
          <div class="bubble-content"><p>Conversation cleared. How can I help with your slides?</p></div>
        </div>
      `;
    }
  });

  // Chat Suggestion Chips
  document.querySelectorAll('.chip-btn').forEach(chip => {
    chip.addEventListener('click', () => {
      const prompt = chip.dataset.prompt;
      const input = document.getElementById('chat-input');
      if (input) {
        input.value = prompt;
        handleSendChatMessage();
      }
    });
  });

  // Slide Generator
  document.getElementById('generate-slide-btn')?.addEventListener('click', handleGenerateSlide);

  // Insert from Preview button (The ONLY time the generated slide is inserted!)
  document.getElementById('insert-preview-btn')?.addEventListener('click', () => {
    if (state.lastGeneratedSlide) {
      insertSlideIntoPowerPoint(state.lastGeneratedSlide);
    } else {
      showToast('Generate a slide first before inserting.');
    }
  });

  // Refine Tab Presets
  document.querySelectorAll('.preset-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const action = btn.dataset.action;
      handleRefineAction(action);
    });
  });

  document.getElementById('refine-execute-btn')?.addEventListener('click', () => {
    const custom = document.getElementById('refine-custom-input')?.value.trim();
    handleRefineAction(custom ? 'custom' : 'concise', custom);
  });

  document.getElementById('copy-refine-btn')?.addEventListener('click', () => {
    const text = document.getElementById('refine-result-text')?.textContent;
    if (text) {
      navigator.clipboard.writeText(text);
      showToast('Copied refined text!');
    }
  });

  // Speaker Notes & Insert Notes Handlers
  document.getElementById('gen-notes-btn')?.addEventListener('click', handleGenerateNotes);

  // Wired up "Add to Slide Notes" button!
  document.getElementById('insert-notes-btn')?.addEventListener('click', () => {
    const notesText = state.lastGeneratedNotes || document.getElementById('notes-result-text')?.textContent;
    if (!notesText) {
      showToast('Generate speaker notes first.');
      return;
    }

    navigator.clipboard.writeText(notesText);
    showToast('Speaker notes copied to clipboard! Paste directly into PowerPoint Notes.');
  });
}

// ============================================================================
// Utilities & Toast
// ============================================================================

function showToast(message, duration = 3000) {
  const toast = document.getElementById('toast');
  if (!toast) return;
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
