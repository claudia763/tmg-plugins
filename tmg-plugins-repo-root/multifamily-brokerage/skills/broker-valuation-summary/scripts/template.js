// ============================================================================
// BROKER VALUATION SUMMARY TEMPLATE
// ============================================================================
// Copy this file into your outputs directory, rename it for the property
// (e.g. oak_creek_valuation.js), then edit ZONE 2 (data) and ZONE 3 (narrative)
// below. Do not touch ZONE 1 — it's the formatting engine that makes every
// document look consistent.
//
// Run with: node <yourfile>.js
// Requires: npm install docx   (run once per session, in the outputs dir)
// ============================================================================

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, Header, Footer, PageNumber, ImageRun
} = require('./node_modules/docx');
const fs = require('fs');

// ============================================================================
// ZONE 1 — INFRASTRUCTURE (do not modify)
// ============================================================================

// The Multifamily Group brand palette (sampled from the official OM):
//   Brand navy  #1B3E6F   |   Brand gold  #FDB714
const NAVY = "1B3E6F";       // TMG brand navy — primary
const DARK_BLUE = "345279";  // coordinated mid-navy (OM secondary tone)
const LIGHT_BLUE = "DCE6F2"; // pale navy tint for subject/callout shading
const GOLD = "FDB714";       // TMG brand gold — accent
const LIGHT_GRAY = "F2F2F2";
const MID_GRAY = "CCCCCC";
const WHITE = "FFFFFF";
const RED = "C00000";
const GREEN = "375623";

// TMG logo (brand navy on white), embedded so the template stays self-contained.
const LOGO_BASE64 = "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IajpPfKzrpoyZM60YdnpSVypM8Awao0NgGdc0HXqqvMqsjRmcPfrLhsxZUx/k0GpN4fU13QXvgsCAgICAgJtRBq48cAX0Ly+QH5O2q8nD77m0uH5uWmgUfn2VtIFjCAA/I7jgr751146fNaECyxmA2chrVbqPJeDTkhtnS/JZrp84gXXzxg5dkiPen5DKYKipruAgICAgEA7kAa3N1BT5x02oMvcqUOvnDo02WEBjYIXWl2+nTFWVeuxW4yTRva5bvqIKaP7cUXO/RzPxr0gGNSratw9OmddP33ENZdd2LtLNgCAAUApRRAJuiAgICAgIND2pAFCyBgb3r/L/82bNGN8QchxodXBC42hyNK8mSNvvnz0BX3a4I6jvsEAAgByMpyP3HX5LbPHZKU5QEMcB0KCLggICAgICLQBIGPNF4hopwQG/I4DQdTmmRHOxHEICAgICAgInAPSwBijlPG0jG0LShkAoD2UukjQJCAgICAg8AOQBgEBAQEBAQGBEMRlv4CAgICAgIAgDQICAgICAgKCNAgICAgICAgI0iAgICAgICAgSIOAgICAgICAIA0CAgICAgICP238P0aZyBHSI9e1AAAAAElFTkSuQmCC";
const LOGO_BUFFER = Buffer.from(LOGO_BASE64, "base64");
const LOGO_W = 220, LOGO_H = 34; // px; preserves the logo's 6.54:1 aspect ratio

const border = { style: BorderStyle.SINGLE, size: 1, color: MID_GRAY };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: WHITE };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function para(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.spaceBefore ?? 80, after: opts.spaceAfter ?? 80 },
    children: [new TextRun({
      text,
      bold: opts.bold || false,
      italics: opts.italics || false,
      size: opts.size || 20,
      color: opts.color || "000000",
      font: opts.font || "Arial",
    })]
  });
}

function richPara(runs, opts = {}) {
  // runs: array of { text, bold, italics, color, size }
  return new Paragraph({
    spacing: { before: opts.spaceBefore ?? 80, after: opts.spaceAfter ?? 80 },
    children: runs.map(r => new TextRun({
      text: r.text,
      bold: r.bold || false,
      italics: r.italics || false,
      size: r.size || 20,
      color: r.color || "000000",
      font: "Arial",
    }))
  });
}

function sectionHeader(text) {
  return new Paragraph({
    spacing: { before: 240, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 4 } },
    children: [new TextRun({
      text: text.toUpperCase(),
      bold: true,
      size: 24,
      color: NAVY,
      font: "Arial",
    })]
  });
}

function subHeader(text) {
  return para(text, { bold: true, size: 20, color: NAVY, spaceBefore: 140, spaceAfter: 60 });
}

function bullet(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    indent: { left: 360, hanging: 200 },
    children: [
      new TextRun({ text: "•  ", bold: false, size: 20, font: "Arial", color: GOLD }),
      new TextRun({ text, bold: opts.bold || false, size: opts.size || 20, font: "Arial", color: opts.color || "000000" })
    ]
  });
}

function spacer(after = 160) {
  return para("", { spaceBefore: 0, spaceAfter: after });
}

function twoColRow(label, value, shaded = false) {
  const fill = shaded ? LIGHT_GRAY : WHITE;
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: 4500, type: WidthType.DXA },
        shading: { fill, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 18, font: "Arial", color: "333333" })] })]
      }),
      new TableCell({
        borders,
        width: { size: 4860, type: WidthType.DXA },
        shading: { fill, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: value, size: 18, font: "Arial", color: "000000" })] })]
      })
    ]
  });
}

// Generic N-column table builder. headers: array of strings. widths: array of DXA ints.
// rows: array of arrays of strings (same length as headers).
function dataTable(headers, widths, rows, opts = {}) {
  const totalWidth = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => new TableCell({
          borders,
          width: { size: widths[i], type: WidthType.DXA },
          shading: { fill: opts.headerFill || NAVY, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 80, right: 80 },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: h, bold: true, size: 16, font: "Arial", color: WHITE })]
          })]
        }))
      }),
      ...rows.map((r, ri) => new TableRow({
        children: r.map((v, i) => new TableCell({
          borders,
          width: { size: widths[i], type: WidthType.DXA },
          shading: { fill: ri % 2 === 0 ? WHITE : LIGHT_GRAY, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 80, right: 80 },
          children: [new Paragraph({
            alignment: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
            children: [new TextRun({ text: v, size: 16, font: "Arial", color: "000000" })]
          })]
        }))
      }))
    ]
  });
}

// Highlighted single full-width row, e.g. "Indicated Value" or "Subject" callouts.
// cells: array of { text, span, fill, color, bold }
function calloutRow(cells, widths) {
  return new TableRow({
    children: cells.map((c, i) => new TableCell({
      borders,
      columnSpan: c.span || 1,
      width: { size: c.span ? widths.slice(i, i + c.span).reduce((a, b) => a + b, 0) : widths[i], type: WidthType.DXA },
      shading: { fill: c.fill || LIGHT_BLUE, type: ShadingType.CLEAR },
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
      children: [new Paragraph({
        alignment: c.align || AlignmentType.LEFT,
        children: [new TextRun({ text: c.text, bold: c.bold !== false, size: 16, font: "Arial", color: c.color || NAVY })]
      })]
    }))
  });
}

function titleBlock(propertyInfo) {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 40, after: 120 },
      children: [new ImageRun({
        type: "png",
        data: LOGO_BUFFER,
        transformation: { width: LOGO_W, height: LOGO_H },
      })]
    }),
    new Paragraph({
      spacing: { before: 0, after: 0 },
      shading: { fill: NAVY, type: ShadingType.CLEAR },
      children: [new TextRun({ text: " ", size: 20, font: "Arial" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 40 },
      children: [new TextRun({ text: "BROKER VALUATION SUMMARY", bold: true, size: 32, font: "Arial", color: NAVY })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: propertyInfo.name, bold: true, size: 48, font: "Arial", color: NAVY })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: propertyInfo.subtitle, size: 22, font: "Arial", color: "555555" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
      children: [new TextRun({ text: propertyInfo.priceLine, bold: true, size: 22, font: "Arial", color: GOLD })]
    }),
    new Paragraph({
      spacing: { before: 0, after: 240 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 16, color: GOLD, space: 1 } },
      children: [new TextRun({ text: "", size: 20 })]
    }),
  ];
}

function metricsTable(kv) {
  // kv: array of [label, value] pairs
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [4500, 4860],
    rows: kv.map((r, i) => twoColRow(r[0], r[1], i % 2 === 0))
  });
}

function headerFooter(propertyInfo) {
  return {
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: GOLD, space: 4 } },
          children: [new TextRun({ text: `CONFIDENTIAL  |  ${propertyInfo.name}  |  Broker Valuation Summary  |  ${propertyInfo.date}`, size: 16, font: "Arial", color: "777777" })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: GOLD, space: 4 } },
          children: [
            new TextRun({ text: "Page ", size: 16, font: "Arial", color: "777777" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "777777" }),
            new TextRun({ text: "  |  CONFIDENTIAL — For Qualified Investors Only", size: 16, font: "Arial", color: "777777" }),
          ]
        })]
      })
    }
  };
}

// ============================================================================
// ZONE 2 — DATA (replace every placeholder value with the real deal data)
// ============================================================================

const propertyInfo = {
  name: "PROPERTY NAME",
  subtitle: "Street Address  |  City, ST Zip  |  ### Units  |  Built ####",
  priceLine: "Recommended Valuation: $XX,XXX,XXX  |  $XX,XXX / Unit  |  Confidential — Month Day, Year",
  date: "Month Day, Year", // used in header/footer/disclaimer
};

const keyMetrics = [
  ["Recommended Valuation", "$XX,XXX,XXX"],
  ["Price Per Unit", "$XX,XXX"],
  ["Price Per SF", "$XXX"],
  ["T-12 Cap Rate", "X.XX%"],
  ["T-3 Cap Rate", "X.XX%"],
  ["Pro Forma Cap Rate", "X.XX%"],
  ["Assumable Loan Balance", "$XX,XXX,XXX"],
  ["Blended Interest Rate", "X.XX%"],
  ["LTV", "XX.XX%"],
  ["Equity Requirement", "~$XX,XXX,XXX"],
  ["Projected IRR", "XX.XX%"],
  ["Equity Multiple", "X.XXx"],
];

// Rent comp table — 7 columns: Property, Units, Year Built, Occupancy, Avg SF, Avg Rent, Avg $/SF
const rentCompHeaders = ["Property", "Units", "Year Built", "Occupancy", "Avg SF", "Avg Rent", "Avg $/SF"];
const rentCompWidths = [1600, 1100, 900, 900, 900, 1100, 1160];
const rentCompRows = [
  ["Comp Name 1", "###", "####", "XX%", "###", "$X,XXX", "$X.XX"],
  ["Comp Name 2", "###", "####", "XX%", "###", "$X,XXX", "$X.XX"],
  ["Comp Name 3", "###", "####", "XX%", "###", "$X,XXX", "$X.XX"],
  ["Comp Average", "###", "####", "XX%", "###", "$X,XXX", "$X.XX"],
];
const rentCompSubjectRow = ["###", "####", "XX%", "###", "$X,XXX", "$X.XX"]; // 6 values, no name column

// Sale comp table — 6 columns: Property, Units, Built, Sale Price, Net Adj., Adj. $/Unit
const saleCompHeaders = ["Property", "Units", "Built", "Sale Price", "Net Adj.", "Adj. $/Unit"];
const saleCompWidths = [1700, 700, 800, 1400, 1000, 1500];
const saleCompRows = [
  ["Sale Comp 1", "###", "####", "$XX,XXX,XXX", "-X.X%", "$XXX,XXX"],
  ["Sale Comp 2", "###", "####", "$XX,XXX,XXX", "-X.X%", "$XXX,XXX"],
  ["Sale Comp 3", "###", "####", "$XX,XXX,XXX", "-X.X%", "$XXX,XXX"],
];
const saleCompIndicatedValue = "$XXX,XXX"; // unadjusted comp-indicated $/unit
const saleCompSubjectPrice = "$XX,XXX"; // subject recommended valuation $/unit

// Agency loan comps — 6 columns: Property, Units, Built, Value/Unit, Cap Rate, Address
const agencyCompRows = [
  ["Agency Comp 1", "###", "####", "$XXX,XXX", "X.XX%", "Address line"],
  ["Agency Comp 2", "###", "####", "$XXX,XXX", "X.XX%", "Address line"],
  ["Agency Comp 3", "###", "####", "$XXX,XXX", "X.XX%", "Address line"],
];
const agencyCompAverage = ["Agency Comp Average", "###", "####", "$XXX,XXX", "X.XX%", "Metro/MSA name"];
const agencyCompSubject = [`SUBJECT — ${propertyInfo.name}`, "###", "####", "$XX,XXX", "X.XX%*", "Subject address"];

// Expense comparison table — 4 columns: Expense Category, Current $/Unit, Normalized $/Unit, Annual Savings
const expenseCompRows = [
  ["Payroll", "$X,XXX", "$X,XXX", "~$XXX,XXX"],
  ["Contract Services", "$XXX", "$XXX", "~$XX,XXX"],
  ["Management Fee", "$XXX", "$XXX", "~$XX,XXX"],
  ["Real Estate Taxes", "$X,XXX", "$X,XXX", "~$XXX,XXX"],
];
const expenseCompTotal = { current: "$X,XXX/unit", normalized: "$X,XXX/unit", annualValue: "~$XXX,XXX/yr = $X.XM Value" };

// Final callout table — three boxes
const calloutBoxes = {
  priceRange: { label: "Suggested Price Range", big: "$XX.XM – $XX.XM", small: "$XX,XXX – $XX,XXX/Unit" },
  returns: { label: "Return Profile (Mid-Point)", big: "XX.XX% IRR", small: "X.XXx Equity Multiple" },
  buyer: { label: "Recommended Target Buyer", big: "Institutional Operator", small: "Loan Assumption / Turnaround Focus" },
};

const agencyFootnote = "* Pro forma cap rate on normalized NOI at recommended valuation";

// ============================================================================
// ZONE 3 — NARRATIVE (write the real story here, in the seller/broker voice)
// ============================================================================
// Every section is an array of Paragraph-producing calls. Use:
//   para("plain text", { ...opts })
//   richPara([{ text, bold, color }, { text, bold, color }, ...])  for mixed-style sentences
//   bullet("text")
//   subHeader("Sub-section Title")
// Write as the listing broker advising ownership — recommendations, disclosure
// strategy, pricing logic — never as a buyer-facing pitch.

const introSection = [
  para("[Opening paragraph: what the property is, recommended valuation, and the headline marketing differentiator (e.g. assumable loan, scale).]"),
  richPara([
    { text: "[Second paragraph: honest acknowledgment of operational distress driving the discount — falling rents, bad debt, elevated opex, no renovation premium. State plainly what the value narrative IS: " },
    { text: "operational upside, assumable debt, institutional scale at a compelling basis", bold: true, color: NAVY },
    { text: ".]" },
  ]),
  para("[Third paragraph: quantify the discount to comps and explain why it's necessary and defensible.]", { spaceAfter: 160 }),
];

const financingSection = [
  para("[Intro: frame the loan as the single most important marketing tool, if applicable.]"),
  bullet("[Loan balance / rate — emphasize savings vs. new origination]"),
  bullet("[LTV — emphasize leverage benefit]"),
  bullet("[Maturity date / remaining term]"),
  bullet("[Debt service / DSCR]"),
  bullet("[Equity requirement and projected returns]"),
  para("[Closing: recommended buyer profile given the loan structure and deal scale.]", { spaceAfter: 160 }),
];

const rentalAnalysisSection = [
  para("[Intro: frame market headwinds as something to disclose proactively, not hide.]"),
  subHeader("[Market] Market Context"),
  bullet("[Supply pipeline data point]"),
  bullet("[Vacancy / rent growth data point]"),
  bullet("[Forward-looking counterweight — why this should improve]"),
  bullet("[Employment / wage data point — long-term demand thesis]"),
  subHeader("Subject Property Rental Performance"),
  para("Rental Comparable Summary", { bold: true, size: 18, color: "555555", spaceBefore: 60, spaceAfter: 80 }),
  // rentCompTable inserted here in build step
  bullet("[Subject vs. comp average rent/occupancy gap — frame as market-wide, not property-specific]"),
  bullet("[Loss to lease commentary]"),
  richPara([
    { text: "[Bad debt figure and trend.] " },
    { text: "The trend is worsening", bold: true, color: RED },
    { text: " — [supporting detail]. We recommend ownership provide buyers with a collections action plan to prevent retrading." },
  ]),
  bullet("[Renovation premium reality check — recommend against leading with it if premium is thin]"),
];

const dealOptimizationSection = [
  richPara([
    { text: "With no supportable renovation premium, the value narrative we are bringing to market centers on " },
    { text: "operational rightsizing as the primary value-creation lever", bold: true, color: NAVY },
    { text: ". [Supporting detail on the biggest expense opportunity.]" },
  ]),
  // expenseCompTable inserted here in build step
  bullet("[Payroll or other top expense line opportunity, framed as buyer-identifiable upside]"),
  bullet("[Total NOI improvement and capitalized value impact]"),
  bullet("[Secondary expense item — flag without formally underwriting, to avoid retrading risk]"),
  bullet("[Pro forma NOI / cap rate achievable through opex normalization alone, no rent growth assumed]"),
];

const saleCompAnalysisSection = [
  para("[Intro: pricing is anchored in traditional sale comps + agency loan comps; both confirm the recommended price reflects distress appropriately while leaving a credible upside story.]"),
  subHeader("Traditional Sale Comparables (Adjusted)"),
  // saleCompTable inserted here in build step
  richPara([
    { text: "[Comp set summary, location, adjusted average.] Our recommended valuation reflects a " },
    { text: "[X]% discount to the comp-indicated value", bold: true, color: RED },
    { text: " — attributable to operational distress, not real estate quality." },
  ]),
  para("[Note on comp scale/size differences and why the recommended price still accounts for them appropriately.]"),
  subHeader("Agency Loan Comparables (Fannie/Freddie MBS)"),
  // agencyTable inserted here in build step
  para(agencyFootnote, { italics: true, size: 16, color: "777777", spaceBefore: 60, spaceAfter: 80 }),
  bullet("[Agency comp average and cap rate vs. subject — frame spread as buyer value-creation opportunity]"),
  bullet("[Most comparable single agency comp — supports operational, not structural, discount narrative]"),
  bullet("[Recommend framing the spread as the core investment pitch in the OM]"),
];

const conclusionSection = [
  para("[Recommendation to proceed with disposition at the recommended valuation; why it's appropriately priced for the environment.]"),
  richPara([
    { text: "Our marketing approach should not position this as a renovation play — [reason]", bold: true, color: NAVY },
    { text: ". Instead, we recommend centering the OM on three pillars: (1) [pillar 1]; (2) [pillar 2]; and (3) [pillar 3]." },
  ]),
  // calloutTable inserted here in build step
];

const keyConsiderations = [
  bullet("[Risk 1: further deterioration buyers may stress-test]"),
  bullet("[Risk 2: most likely retrading lever — recommend a specific mitigant]"),
  bullet("[Risk 3: e.g. loan maturity timeline — recommend proactive OM disclosure]"),
];

const disclaimer = "This valuation summary has been prepared by the listing broker on behalf of ownership for internal disposition planning purposes. All information is derived from owner-provided operating statements and market data. This document is confidential and intended solely for ownership and authorized representatives.";

// ============================================================================
// DOCUMENT ASSEMBLY (rarely needs changes — adjust only to add/remove sections)
// ============================================================================

function buildRentCompTable() {
  const table = dataTable(rentCompHeaders, rentCompWidths, rentCompRows);
  // Append subject row with navy "SUBJECT" label cell
  table.root.push(new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: rentCompWidths[0], type: WidthType.DXA },
        shading: { fill: NAVY, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        children: [new Paragraph({ children: [new TextRun({ text: "SUBJECT", bold: true, size: 16, font: "Arial", color: WHITE })] })]
      }),
      ...rentCompSubjectRow.map((v, i) => new TableCell({
        borders,
        width: { size: rentCompWidths[i + 1], type: WidthType.DXA },
        shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: v, bold: true, size: 16, font: "Arial", color: NAVY })] })]
      }))
    ]
  }));
  return table;
}

function buildSaleCompTable() {
  const table = dataTable(saleCompHeaders, saleCompWidths, saleCompRows);
  table.root.push(calloutRow([
    { text: "Unadjusted Comp-Indicated Value/Unit (Equal Weighting)", span: 5, fill: LIGHT_BLUE, color: NAVY },
    { text: saleCompIndicatedValue, fill: LIGHT_BLUE, color: NAVY, align: AlignmentType.CENTER },
  ], saleCompWidths));
  table.root.push(calloutRow([
    { text: "Subject — Recommended Valuation/Unit (Distress Discount)", span: 5, fill: NAVY, color: WHITE },
    { text: saleCompSubjectPrice, fill: GOLD, color: WHITE, align: AlignmentType.CENTER },
  ], saleCompWidths));
  return table;
}

function buildAgencyTable() {
  const headers = ["Property", "Units", "Built", "Value/Unit", "Cap Rate", "Address"];
  const widths = [2200, 700, 700, 1200, 1300, 2660];
  const table = dataTable(headers, widths, agencyCompRows, { headerFill: DARK_BLUE });
  table.root.push(calloutRow(
    agencyCompAverage.map((v, i) => ({ text: v, fill: i === 3 ? GOLD : DARK_BLUE, color: WHITE, align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER })),
    widths
  ));
  table.root.push(calloutRow(
    agencyCompSubject.map((v, i) => ({ text: v, fill: i === 0 ? NAVY : (i === 3 ? GOLD : LIGHT_BLUE), color: i === 0 || i === 3 ? WHITE : NAVY, align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER })),
    widths
  ));
  return table;
}

function buildExpenseTable() {
  const headers = ["Expense Category", "Current $/Unit", "Normalized $/Unit", "Annual Savings"];
  const widths = [3000, 2000, 2000, 2360];
  const table = dataTable(headers, widths, expenseCompRows);
  table.root.push(calloutRow([
    { text: "NET OPEX IMPROVEMENT", fill: LIGHT_BLUE, color: NAVY },
    { text: `${expenseCompTotal.current} → ${expenseCompTotal.normalized}`, span: 2, fill: LIGHT_BLUE, color: NAVY, align: AlignmentType.CENTER },
    { text: expenseCompTotal.annualValue, fill: GOLD, color: WHITE, align: AlignmentType.CENTER },
  ], widths));
  return table;
}

function buildCalloutTable() {
  const boxes = [calloutBoxes.priceRange, calloutBoxes.returns, calloutBoxes.buyer];
  const fills = [NAVY, DARK_BLUE, NAVY];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [3120, 3120, 3120],
    rows: [new TableRow({
      children: boxes.map((b, i) => new TableCell({
        borders,
        width: { size: 3120, type: WidthType.DXA },
        shading: { fill: fills[i], type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 160, right: 160 },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: b.label, bold: true, size: 18, font: "Arial", color: WHITE })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 40 }, children: [new TextRun({ text: b.big, bold: true, size: 24, font: "Arial", color: GOLD })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 20 }, children: [new TextRun({ text: b.small, size: 18, font: "Arial", color: WHITE })] }),
        ]
      }))
    })]
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 }, // US Letter — required, default is A4
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    ...headerFooter(propertyInfo),
    children: [
      ...titleBlock(propertyInfo),

      para("KEY METRICS AT RECOMMENDED VALUATION", { bold: true, size: 18, color: "555555", spaceBefore: 0, spaceAfter: 80 }),
      metricsTable(keyMetrics),
      spacer(160),

      sectionHeader("Introduction"),
      ...introSection,

      sectionHeader("Financing"),
      ...financingSection,

      sectionHeader("Rental Analysis"),
      ...rentalAnalysisSection,
      buildRentCompTable(),
      spacer(160),

      sectionHeader("Deal Optimization / Value-Add"),
      ...dealOptimizationSection.slice(0, 1),
      buildExpenseTable(),
      spacer(80),
      ...dealOptimizationSection.slice(1),
      spacer(160),

      sectionHeader("Sale Comparable Analysis"),
      ...saleCompAnalysisSection.slice(0, 2),
      buildSaleCompTable(),
      spacer(80),
      ...saleCompAnalysisSection.slice(2, 5),
      buildAgencyTable(),
      ...saleCompAnalysisSection.slice(5),
      spacer(160),

      sectionHeader("Conclusion"),
      ...conclusionSection,
      buildCalloutTable(),
      spacer(120),

      para("Key Considerations for Seller / Pricing Risk Factors:", { bold: true, spaceBefore: 80, spaceAfter: 80 }),
      ...keyConsiderations,
      spacer(120),

      new Paragraph({
        spacing: { before: 60, after: 60 },
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: MID_GRAY, space: 4 } },
        children: [new TextRun({ text: disclaimer, italics: true, size: 16, font: "Arial", color: "888888" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  const outName = `${propertyInfo.name.replace(/\s+/g, '_')}_Broker_Valuation_Summary.docx`;
  fs.writeFileSync(outName, buffer);
  console.log('Done:', outName);
});
