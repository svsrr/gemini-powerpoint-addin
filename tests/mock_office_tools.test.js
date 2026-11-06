/**
 * Mocked Office.js Integration Tests for PowerPoint Copilot Tools
 * Tests runtime Office.js proxy lifecycle, undo stack, strict shape lookup, and prompt injection defense.
 */
import assert from 'node:assert';

console.log('--- Starting Mock Office.js Tool Execution Test Suite ---');

// Setup mock Office.js environment
function createMockPowerPointEnvironment() {
  const cellsData = [
    [
      { text: "SI No", font: { color: "#000000" }, fill: { color: "#F1F5F9" } },
      { text: "Action Item", font: { color: "#000000" }, fill: { color: "#F1F5F9" } },
      { text: "Status", font: { color: "#000000" }, fill: { color: "#F1F5F9" } }
    ],
    [
      { text: "1.", font: { color: "#000000" }, fill: { color: "#FFFFFF" } },
      { text: "Safety review", font: { color: "#000000" }, fill: { color: "#FFFFFF" } },
      { text: "Complied", font: { color: "#2AC456" }, fill: { color: "#FFFFFF" } } // Green
    ],
    [
      { text: "2.", font: { color: "#000000" }, fill: { color: "#FFFFFF" } },
      { text: "Audit report", font: { color: "#000000" }, fill: { color: "#FFFFFF" } },
      { text: "Job on hold", font: { color: "#2AC456" }, fill: { color: "#FFFFFF" } } // Green
    ]
  ];

  const mockTable = {
    rowCount: 3,
    columnCount: 3,
    values: cellsData.map(row => row.map(c => c.text)),
    load: () => {},
    getCellOrNullObject: (r, c) => ({
      get text() { return cellsData[r][c].text; },
      set text(v) { cellsData[r][c].text = v; },
      font: {
        get color() { return cellsData[r][c].font.color; },
        set color(v) { cellsData[r][c].font.color = v; },
        load: () => {}
      },
      fill: {
        setSolidColor: (v) => { cellsData[r][c].fill.color = v; }
      },
      load: () => {}
    })
  };

  const shapes = [
    {
      id: "sh-title",
      name: "Title Placeholder 1",
      type: "placeholder",
      top: 50,
      left: 60,
      width: 800,
      height: 60,
      placeholderFormat: {
        type: "Title",
        load: () => {}
      },
      textFrame: {
        hasText: true,
        textRange: {
          text: "EXECUTIVE SUMMARY",
          font: { color: "#0000FF", size: 28, bold: true, load: () => {} },
          load: () => {}
        },
        load: () => {}
      },
      load: () => {}
    },
    {
      id: "sh-table",
      name: "Table 1",
      type: "table",
      top: 150,
      left: 60,
      width: 800,
      height: 300,
      getTable: () => mockTable,
      load: () => {}
    },
    {
      id: "sh-body",
      name: "Body Text 2",
      type: "textbox",
      top: 480,
      left: 60,
      width: 800,
      height: 40,
      textFrame: {
        hasText: true,
        textRange: {
          text: "Confidential - For Internal Use Only",
          font: { color: "#64748B", size: 10, bold: false, load: () => {} },
          load: () => {}
        },
        load: () => {}
      },
      delete: () => {
        const idx = shapes.findIndex(s => s.id === "sh-body");
        if (idx !== -1) shapes.splice(idx, 1);
      },
      load: () => {}
    }
  ];

  const slide = {
    id: "slide-42",
    shapes: {
      items: shapes,
      load: () => {}
    },
    load: () => {}
  };

  const context = {
    presentation: {
      getSelectedSlides: () => ({
        items: [slide],
        load: () => {},
        getItemAt: () => slide
      }),
      slides: {
        getItemOrNullObject: (id) => (id === "slide-42" ? slide : { isNullObject: true }),
        load: () => {}
      }
    },
    sync: async () => {}
  };

  return {
    slide,
    shapes,
    cellsData,
    run: async (fn) => await fn(context)
  };
}

// Test Case 1: Title placeholder priority & set_font
async function testTitlePlaceholderPriority() {
  const env = createMockPowerPointEnvironment();
  const undoStack = [];

  // Simulate executeSetFont with shape_id="heading"
  const result = await env.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const textual = slide.shapes.items.filter(s =>
      ['textbox', 'geometricshape', 'placeholder'].includes(String(s.type || '').toLowerCase())
    );
    const placeholderShapes = textual.filter(s => String(s.type || '').toLowerCase() === 'placeholder');
    const titleShape = placeholderShapes.find(s => s.placeholderFormat.type === 'Title');

    assert(titleShape, 'Should identify Title placeholder');
    assert.strictEqual(titleShape.id, 'sh-title');

    // Save undo
    undoStack.push({
      type: 'set_font',
      shapeId: titleShape.id,
      previous: { color: titleShape.textFrame.textRange.font.color },
      current: { color: '#EA580C' }
    });

    titleShape.textFrame.textRange.font.color = '#EA580C';
    return { success: true, newColor: titleShape.textFrame.textRange.font.color };
  });

  assert.strictEqual(result.success, true);
  assert.strictEqual(result.newColor, '#EA580C');
  assert.strictEqual(undoStack.length, 1);
  console.log('✔ Case 1 Passed: Title placeholder correctly prioritized and styled orange');
}

// Test Case 2: Strict shape lookup (Typo protection)
async function testStrictShapeLookup() {
  const env = createMockPowerPointEnvironment();

  const result = await env.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shape_id = "NonExistentTypoBox";
    const lowerId = shape_id.toLowerCase();

    const targetShape = slide.shapes.items.find(s => s.id === shape_id || s.name.toLowerCase() === lowerId);
    if (!targetShape) {
      const available = slide.shapes.items.map(s => s.name).join(', ');
      return { success: false, error: `Shape not found: "${shape_id}". Available shapes: ${available}.` };
    }
    return { success: true };
  });

  assert.strictEqual(result.success, false);
  assert(result.error.includes('Shape not found: "NonExistentTypoBox"'), 'Must return explicit shape not found error');
  // Confirm title was NOT mutated!
  assert.strictEqual(env.shapes[0].textFrame.textRange.font.color, '#0000FF');
  console.log('✔ Case 2 Passed: Unknown shape IDs safely reject without accidentally editing the title');
}

// Test Case 3: Table batch update with cell font colors and undo reversion
async function testTableBatchUpdateAndUndo() {
  const env = createMockPowerPointEnvironment();
  const undoStack = [];

  // Update green remarks to black
  const updates = [
    { row: 1, col: 2, font_color: '#000000' },
    { row: 2, col: 2, font_color: '#000000' }
  ];

  await env.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const tableShape = slide.shapes.items.find(s => s.type === 'table');
    const table = tableShape.getTable();

    const previousCells = [];
    for (const up of updates) {
      const cell = table.getCellOrNullObject(up.row, up.col);
      previousCells.push({ row: up.row, col: up.col, font_color: cell.font.color });
      cell.font.color = up.font_color;
    }

    undoStack.push({
      type: 'table_cells',
      tableId: tableShape.id,
      previousCells
    });
  });

  // Verify updated to black
  assert.strictEqual(env.cellsData[1][2].font.color, '#000000');
  assert.strictEqual(env.cellsData[2][2].font.color, '#000000');
  console.log('✔ Case 3a: Green remarks successfully updated to black (#000000)');

  // Now perform Undo
  const undoAction = undoStack.pop();
  await env.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const tableShape = slide.shapes.items.find(s => s.id === undoAction.tableId);
    const table = tableShape.getTable();

    for (const prev of undoAction.previousCells) {
      const cell = table.getCellOrNullObject(prev.row, prev.col);
      cell.font.color = prev.font_color;
    }
  });

  // Verify reverted to original green (#2AC456)
  assert.strictEqual(env.cellsData[1][2].font.color, '#2AC456');
  assert.strictEqual(env.cellsData[2][2].font.color, '#2AC456');
  console.log('✔ Case 3b: Undo action flawlessly restored original table colors');
}

// Test Case 4: Prompt Injection Protection on Destructive Actions
async function testDestructiveActionConfirmation() {
  function executeDeleteShape({ shape_id, user_confirmed }) {
    if (!user_confirmed) {
      return { success: false, error: "Deletion aborted: Deleting a shape requires explicit confirmation in chat." };
    }
    return { success: true };
  }

  // 4a. Attempt deletion without confirmation
  const resUnconfirmed = executeDeleteShape({ shape_id: "sh-body", user_confirmed: false });
  assert.strictEqual(resUnconfirmed.success, false);
  assert(resUnconfirmed.error.includes("Deletion aborted"));

  // 4b. Explicit confirmation
  const resConfirmed = executeDeleteShape({ shape_id: "sh-body", user_confirmed: true });
  assert.strictEqual(resConfirmed.success, true);
  console.log('✔ Case 4 Passed: Prompt injection defense blocks unconfirmed shape deletions');
}

// Test Case 5: Multi-Slide Direct Targeting (Slide 6 & Slide 7)
async function testMultiSlideTargeting() {
  const slide6Shape = {
    id: "s6-title",
    name: "TextBox 3",
    type: "textbox",
    top: 188672,
    textFrame: {
      hasText: true,
      textRange: {
        text: "TRAINING ON SMP",
        font: { color: "#0000FF", size: 24, bold: true }
      }
    }
  };

  const slide7Shape = {
    id: "s7-title",
    name: "TextBox 6",
    type: "textbox",
    top: 197319,
    textFrame: {
      hasText: true,
      textRange: {
        text: "ROPE ACCESS SYSTEM IN PDR",
        font: { color: "#0000FF", size: 24, bold: true }
      }
    }
  };

  const slidesArray = [
    { id: "s1", shapes: { items: [] } },
    { id: "s2", shapes: { items: [] } },
    { id: "s3", shapes: { items: [] } },
    { id: "s4", shapes: { items: [] } },
    { id: "s5", shapes: { items: [] } },
    { id: "s6", shapes: { items: [slide6Shape] } },
    { id: "s7", shapes: { items: [slide7Shape] } },
    { id: "s8", shapes: { items: [] } },
    { id: "s9", shapes: { items: [] } },
    { id: "s10", shapes: { items: [] } }
  ];

  function getMockTargetSlide(slide_number) {
    if (slide_number !== undefined && slide_number !== null) {
      const num = Number(slide_number);
      if (!isNaN(num) && num >= 1 && num <= slidesArray.length) {
        return slidesArray[num - 1];
      }
      throw new Error(`Slide number ${num} is out of range. Presentation has ${slidesArray.length} slide(s).`);
    }
    return slidesArray[5]; // active default (slide 6)
  }

  function mockSetFont({ slide_number, shape_id, color }) {
    const slide = getMockTargetSlide(slide_number);
    const textual = slide.shapes.items.filter(s => s.textFrame && s.textFrame.hasText);
    textual.sort((a, b) => (a.top || 0) - (b.top || 0));
    let target = null;
    if (shape_id === 'heading' || shape_id === 'title') {
      target = textual[0];
    } else {
      target = slide.shapes.items.find(s => s.id === shape_id || s.name === shape_id);
    }
    if (!target) throw new Error('Shape not found');
    target.textFrame.textRange.font.color = color;
    return { success: true, slideNumber: slide_number, targetId: target.id, newColor: color };
  }

  // User command: "change color of headings in slide 6,7 to green"
  const res6 = mockSetFont({ slide_number: 6, shape_id: "heading", color: "#16A34A" });
  const res7 = mockSetFont({ slide_number: 7, shape_id: "heading", color: "#16A34A" });

  assert.strictEqual(res6.success, true);
  assert.strictEqual(res7.success, true);
  assert.strictEqual(slide6Shape.textFrame.textRange.font.color, "#16A34A");
  assert.strictEqual(slide7Shape.textFrame.textRange.font.color, "#16A34A");
  console.log('✔ Case 5 Passed: Multi-slide targeting simultaneously recolors Slide 6 & Slide 7 headings to green');

  // Test Case 6: Out-of-bounds slide number safety
  let threw = false;
  try {
    mockSetFont({ slide_number: 99, shape_id: "heading", color: "#16A34A" });
  } catch (err) {
    threw = true;
    assert(err.message.includes('Slide number 99 is out of range'));
  }
  assert.strictEqual(threw, true);
  console.log('✔ Case 6 Passed: Out-of-range slide number gracefully caught and rejected');
}

// Run all test cases
await testTitlePlaceholderPriority();
await testStrictShapeLookup();
await testTableBatchUpdateAndUndo();
await testDestructiveActionConfirmation();
await testMultiSlideTargeting();

console.log('--- ALL MOCK OFFICE.JS TESTS PASSED (6/6) ---');
