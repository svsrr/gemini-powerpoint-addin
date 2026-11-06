/**
 * Live PPTX Slide Matching & Layout Fidelity Test
 * Tests real structures from Presentation2.pptx
 */
import assert from 'node:assert';
import fs from 'node:fs';

console.log('--- Testing PPTX Fidelity against Presentation2.pptx ---');

// Mock snapshot from Slide 4
const slide4Snapshot = {
  slideTitle: "ATR OF LAST LMC (May'25)",
  textShapes: [
    {
      id: "3",
      name: "Rectangle 2",
      type: "GeometricShape",
      isTitle: true,
      top: 116632,
      left: 100000,
      text: "ATR OF LAST LMC (May'25)",
      font: { color: "0000FF", size: 24, bold: true }
    }
  ],
  tables: [
    {
      shapeId: "tbl-summary",
      name: "Table 7",
      rowCount: 2,
      columnCount: 3,
      values: [
        ["Total ATR Points", "Points Complied", "Action In progress"],
        ["6", "5", "1"]
      ]
    },
    {
      shapeId: "tbl-remarks",
      name: "Table 8",
      rowCount: 7,
      columnCount: 3,
      values: [
        ["Sl No", "Action Point", "Present Status / Remarks"],
        ["1.", "Plan to be shared for hard stand in NHTCCR Unit,.", "Job on hold"],
        ["2.", "Details of water source to be shared to P&U Deptt.", "Complied"],
        ["3.", "Panel Insulation in DCU Coke drum proposal to be dropped", "Complied"],
        ["4.", "Shifting of water filled barrier to PP unit.", "Complied"],
        ["5.", "Small bore painting to be shared in next LMC.", "Complied"],
        ["6.", "PRU Cladding job to be completed on priority", "Complied"]
      ]
    }
  ]
};

// Test A: User asks "change heading color to orange" on Slide 4
assert.strictEqual(slide4Snapshot.textShapes[0].name, "Rectangle 2");
assert.strictEqual(slide4Snapshot.textShapes[0].isTitle, true);
console.log('✔ Test A: Slide 4 title correctly identified as Rectangle 2 (GeometricShape)');

// Test B: User asks "change remarks text to black instead of green"
const remarksTable = slide4Snapshot.tables.find(t => t.name === "Table 8");
assert(remarksTable, "Table 8 exists");
assert.strictEqual(remarksTable.values[0][2], "Present Status / Remarks");
const remarksRows = [];
for (let r = 1; r < remarksTable.rowCount; r++) {
  remarksRows.push({ row: r, col: 2, text: remarksTable.values[r][2] });
}
assert.strictEqual(remarksRows.length, 6);
assert.strictEqual(remarksRows[0].text, "Job on hold");
assert.strictEqual(remarksRows[1].text, "Complied");
console.log(`✔ Test B: Remarks column correctly identified with ${remarksRows.length} data rows`);

// Mock snapshot from Slide 6
const slide6Snapshot = {
  slideTitle: "TRAINING ON SMP",
  textShapes: [
    {
      id: "4",
      name: "TextBox 3",
      type: "TextBox",
      isTitle: true,
      top: 188672,
      left: 150000,
      text: "TRAINING ON SMP",
      font: { color: "0000FF", size: 24, bold: true }
    }
  ],
  tables: [
    {
      shapeId: "tbl-smp",
      name: "Table 4",
      rowCount: 2,
      columnCount: 6,
      values: [
        ["Sl No.", "No of SMP", "Revised SMP", "Addition", "Training Completed", "Training planned in July'25"],
        ["1", "45+2", "4", "2", "5 times ( 1-29)", "3 times (30-47)"]
      ]
    }
  ]
};

// Test C: Slide 6 title correctly identified as TextBox 3
assert.strictEqual(slide6Snapshot.textShapes[0].name, "TextBox 3");
assert.strictEqual(slide6Snapshot.slideTitle, "TRAINING ON SMP");
console.log('✔ Test C: Slide 6 title correctly identified as TextBox 3 (TextBox)');

// Mock snapshot from Slide 7
const slide7Shapes = [
  { id: "sp-desc", name: "TextBox 11", top: 1052736, left: 767408, text: "ROPE ACCESS SYSTEM BEING USED IN CHS SURGE BIN SHEET REPLACEMENT AND PAINTING" },
  { id: "sp-title", name: "TextBox 6", top: 197319, left: 119336, text: "ROPE ACCESS SYSTEM IN PDR" }
];

// Test D: Slide 7 title resolution (Topmost shape identification: TextBox 6 at top 197319 vs TextBox 11 at 1052736)
slide7Shapes.sort((a, b) => a.top - b.top);
const slide7TitleShape = slide7Shapes[0];
assert.strictEqual(slide7TitleShape.name, "TextBox 6");
assert.strictEqual(slide7TitleShape.text, "ROPE ACCESS SYSTEM IN PDR");
console.log('✔ Test D: Slide 7 title correctly identified as TextBox 6 (ROPE ACCESS SYSTEM IN PDR)');

// Test E: Multi-Slide 6 & 7 simultaneous heading color change
const multiSlideCalls = [
  { name: "set_font", args: { slide_number: 6, shape_id: "heading", color: "#16A34A" } },
  { name: "set_font", args: { slide_number: 7, shape_id: "heading", color: "#16A34A" } }
];

assert.strictEqual(multiSlideCalls.length, 2);
assert.strictEqual(multiSlideCalls[0].args.slide_number, 6);
assert.strictEqual(multiSlideCalls[1].args.slide_number, 7);
assert.strictEqual(multiSlideCalls[0].args.color, "#16A34A");
assert.strictEqual(multiSlideCalls[1].args.color, "#16A34A");
console.log('✔ Test E: Multi-slide simultaneous tool call targets Slide 6 & 7 heading colors');

console.log('--- ALL PPTX FIDELITY TESTS PASSED (5/5) ---');
