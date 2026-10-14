// Build the executive deck from deck_data.json (exported from the recalculated workbook).
// Run: NODE_PATH=<dir with pptxgenjs> node deck/build_deck.js deck/deck_data.json deck/Rippling_NB_Segmentation_Staffing.pptx <apply_theme.js>
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const [, , DATA, OUT, THEME_JS] = process.argv;
const D = JSON.parse(fs.readFileSync(DATA, "utf8"));
const { applyTheme } = require(THEME_JS);

// ---------------------------------------------------------------- palette & helpers
const THEME = {
  name: "RevOps Case", headFontFace: "Calibri", bodyFontFace: "Calibri",
  colors: { dk1: "1E1B24", lt1: "FFFFFF", dk2: "4B1D3F", lt2: "F3EFF2", accent1: "4B1D3F", accent2: "C98A2B",
            accent3: "8C6A80", accent4: "BFB3BC", accent5: "2E7D6B", accent6: "B23A48", hlink: "4B1D3F", folHlink: "8C6A80" },
};
const P = { ink: "1E1B24", plum: "4B1D3F", amber: "C98A2B", mauve: "8C6A80", grey: "6B6470", light: "F3EFF2",
            line: "D9D2D7", white: "FFFFFF", teal: "2E7D6B", red: "B23A48", soft: "BFB3BC" };
const m = (x, d = 2) => `$${(x / 1e6).toFixed(d)}M`;
const k = (x) => `$${(x / 1e3).toFixed(1)}K`;
const usd = (x) => `$${Math.round(x).toLocaleString("en-US")}`;
const pct = (x, d = 1) => `${(x * 100).toFixed(d)}%`;
const pts = (x) => `${x >= 0 ? "+" : "−"}${Math.abs(x * 100).toFixed(1)} pts`;
const n0 = (x) => Math.round(x).toLocaleString("en-US");
const SEG = ["1–50", "51–100", "101–500", "501+"];
const SEGN = ["Velocity", "Emerging MM", "Core MM", "Enterprise"];
const S = D.staff, EB = D.early_vs_base, F = D.seg_funnel;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "New-Business Segmentation and Sales Staffing";

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: P.white },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 12.33, h: 0.95, fontSize: 26, bold: true,
        color: P.plum, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.4, y: 7.05, w: 0.5, h: 0.3, fontSize: 10, color: P.grey, align: "right" },
});
pres.defineSlideMaster({
  title: "DARK",
  background: { color: P.plum },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.3, w: 11.7, h: 1.6, fontSize: 40, bold: true,
        color: P.white, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.1, w: 11.7, h: 1.2, fontSize: 18,
        color: "E9DDE5", valign: "top", margin: 0 }, text: "" } },
  ],
});

function footer(slide, text) {
  slide.addText(text, { x: 0.5, y: 7.0, w: 11.6, h: 0.35, fontSize: 10, color: P.grey, isTextBox: true, margin: 0, valign: "middle",
                        objectName: "Source note" });
}
function sub(slide, text, y = 1.2) {
  slide.addText(text, { x: 0.5, y, w: 12.33, h: 0.4, fontSize: 15, color: P.grey, isTextBox: true, margin: 0, objectName: "Subtitle" });
}
function badge(slide, num, x, y, size = 0.5) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: size, h: size, fill: { color: P.plum }, line: { color: P.plum }, objectName: `Badge ${num}` });
  slide.addText(String(num), { x, y, w: size, h: size, fontSize: 16, bold: true, color: P.white, align: "center", valign: "middle",
                               isTextBox: true, margin: 0 });
}
function card(slide, x, y, w, h, fill = P.light, name = "Card") {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, objectName: name });
}
function stat(slide, x, y, w, big, label, color = P.plum) {
  slide.addText(big, { x, y, w, h: 0.6, fontSize: 30, bold: true, color, isTextBox: true, margin: 0 });
  slide.addText(label, { x, y: y + 0.62, w, h: 0.6, fontSize: 13, color: P.ink, isTextBox: true, margin: 0, valign: "top" });
}
const chartText = { catAxisLabelColor: P.grey, valAxisLabelColor: P.grey, catAxisLabelFontSize: 11, valAxisLabelFontSize: 11,
  catAxisLabelFontFace: "Calibri", valAxisLabelFontFace: "Calibri", dataLabelFontFace: "Calibri", dataLabelFontSize: 11,
  dataLabelColor: P.ink, legendFontFace: "Calibri", legendFontSize: 11, titleFontFace: "Calibri", titleFontSize: 13, titleColor: P.ink };
const SRC = "Source: case dataset (fictional), Feb-2021 to Jan-2022. Rates use summed numerators and denominators; same-month ratios, not cohorts.";
const SRC_STAFF = "Staffing assumptions are illustrative, not company data. Headcount = required capacity, not net-new hires (current headcount unknown).";

// ---------------------------------------------------------------- derived values
const totF = F[4];
const base12 = S.arr_y[4], up12 = S.u_arr_y[4];
const l2dRel = EB.bd[4] / EB.ed[4] - 1, d2oRel = EB.bq[4] / EB.eq[4] - 1, o2wRel = EB.bw[4] / EB.ew[4] - 1;
const upmarketShare = F[2].arr_share + F[3].arr_share;
const bk = D.bucket_funnel; // 1-10, 11-50, 51-100, 101-500, 501+, total
const optA_spread = bk[2].asp / bk[0].asp;
const optB_spread = bk[2].l2w / bk[3].l2w;
const aeAt = (f) => S.demos.slice(0, 4).reduce((t, d, i) => t + Math.ceil(d / (S.ae_cap[i] * f * S.util[i])), 0);
const hcSdr = S.h_sdr[4], hcAe = S.h_ae[4], hcSc = D.sc_tot, hc = D.hc_tot;

// ================================================================ 0. Cover
pres.addSection({ title: "Main" });
let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Main" });
s.addText("New-business segmentation and sales staffing", { placeholder: "title" });
s.addText("Revenue Strategy & Operations case study  |  Fictional dataset, Feb-2021 to Jan-2022", { placeholder: "body" });
s.addNotes("Introduce the goal: show how I'd read the business, segment prospects, and staff SDRs, AEs and SCs to grow new-logo ARR. Seven slides, then appendix.");

// ================================================================ 1. Executive summary
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Main" });
s.addText("Growth is volume-led; the next ARR gains come from fixing first touch and staffing by segment", { placeholder: "title" });
sub(s, `Monthly new ARR grew ${pct(EB.ga[4], 0)} (Feb–Jul to Aug–Oct 2021 average) to a ${m(base12, 1)} annualized run-rate. Three findings shape the plan.`);
const findings = [
  ["Leads outran demos", `Leads +${pct(EB.gl[4], 0)}, but lead→demo fell ${pct(EB.ed[4])} → ${pct(EB.bd[4])}. A potential top-of-funnel bottleneck; cause not yet proven.`],
  ["51–100 has the best revenue yield", `${usd(F[1].arr_lead)} ARR per lead, highest of any segment: small-business conversion at ${(F[1].asp / F[0].asp).toFixed(1)}× the deal size.`],
  ["Growth is moving upmarket", `101+ employees = ${pct(upmarketShare, 0)} of ARR and growing fastest, but lead→win falls to ${pct(F[2].l2w)} and ${pct(F[3].l2w)}.`],
];
findings.forEach(([h, b], i) => {
  const y = 1.9 + i * 1.5;
  card(s, 0.5, y, 7.3, 1.3, P.light, `Finding ${i + 1}`);
  badge(s, i + 1, 0.75, y + 0.4);
  s.addText(h, { x: 1.45, y: y + 0.15, w: 6.2, h: 0.4, fontSize: 17, bold: true, color: P.plum, isTextBox: true, margin: 0 });
  s.addText(b, { x: 1.45, y: y + 0.55, w: 6.2, h: 0.7, fontSize: 14, color: P.ink, isTextBox: true, margin: 0, valign: "top" });
});
card(s, 8.1, 1.9, 4.73, 4.3, P.plum, "Recommendation panel");
s.addText("Recommendation", { x: 8.4, y: 2.05, w: 4.2, h: 0.4, fontSize: 17, bold: true, color: P.white, isTextBox: true, margin: 0 });
s.addText([
  { text: "Four segments, three motions: Velocity 1–50, Mid-Market 51–100 and 101–500, Enterprise 501+", options: { bullet: true, breakLine: true } },
  { text: `${hc} roles of required capacity: ${hcSdr} SDR, ${hcAe} AE, ${hcSc} SC. Hire only the gap to current headcount`, options: { bullet: true, breakLine: true } },
  { text: "Diagnose and fix lead→demo before buying more pipeline", options: { bullet: true, breakLine: true } },
  { text: "SC support upmarket; hold Enterprise expansion until win rates improve", options: { bullet: true } },
], { x: 8.4, y: 2.5, w: 4.2, h: 3.6, fontSize: 14, color: P.white, isTextBox: true, margin: 0, valign: "top", paraSpaceAfter: 8 });
footer(s, SRC);
s.addNotes(`Headline: the business is growing, mostly because more leads come in. Conversion at the first step is slipping and the mix is moving to larger, harder deals. That shapes three things: fix first touch, segment where the data breaks, and staff each segment for its motion. Baseline run-rate ${m(base12)} = Aug-Oct average x 12.`);

// ================================================================ 2. Finding 1
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Main" });
s.addText(`Leads grew ${pct(EB.gl[4], 0)}, but a smaller share reached a demo: a potential top-of-funnel bottleneck`, { placeholder: "title" });
sub(s, "Monthly leads vs demos, all segments. Demos did not keep pace as lead volume climbed.");
s.addChart(pres.charts.LINE, [
  { name: "Leads", labels: D.monthly.map((r) => r.month), values: D.monthly.map((r) => r.leads) },
  { name: "Demos", labels: D.monthly.map((r) => r.month), values: D.monthly.map((r) => r.demos) },
], { x: 0.5, y: 1.7, w: 7.2, h: 4.9, chartColors: [P.soft, P.plum], lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 6,
     showLegend: true, legendPos: "t", valGridLine: { color: "E6E1E5", size: 0.75 }, catGridLine: { style: "none" },
     valAxisLabelFormatCode: "#,##0", ...chartText, objectName: "Leads vs demos chart" });
s.addText("Nov–Dec dip seen once; not treated as confirmed seasonality", { x: 0.6, y: 6.55, w: 7.0, h: 0.3, fontSize: 11, italic: true,
  color: P.grey, isTextBox: true, margin: 0 });
// right: stage table
s.addText("Average month: Feb–Jul vs Aug–Oct 2021", { x: 8.0, y: 1.7, w: 4.83, h: 0.35, fontSize: 14, bold: true, color: P.ink, isTextBox: true, margin: 0 });
const hdr = (t) => ({ text: t, options: { bold: true, color: P.white, fill: { color: P.plum } } });
const stageRows = [
  [hdr("Stage"), hdr("Feb–Jul"), hdr("Aug–Oct"), hdr("Rel. change")],
  ["Lead → demo", pct(EB.ed[4]), pct(EB.bd[4]), { text: pct(l2dRel, 0), options: { bold: true, color: P.red } }],
  ["Demo → opp", pct(EB.eq[4]), pct(EB.bq[4]), pct(d2oRel, 0)],
  ["Opp → win", pct(EB.ew[4]), pct(EB.bw[4]), pct(o2wRel, 0)],
];
s.addTable(stageRows, { x: 8.0, y: 2.1, w: 4.83, colW: [1.53, 1.1, 1.1, 1.1], fontSize: 13, color: P.ink, border: { type: "solid", pt: 0.5, color: P.line },
  align: "center", valign: "middle", rowH: 0.38, objectName: "Stage table" });
s.addText("Lead→demo change by segment", { x: 8.0, y: 3.85, w: 4.83, h: 0.35, fontSize: 14, bold: true, color: P.ink, isTextBox: true, margin: 0 });
s.addTable([SEG.map((t) => hdr(t)), EB.cd.slice(0, 4).map((v) => ({ text: pts(v), options: { color: v < -0.02 ? P.red : P.ink, bold: v < -0.02 } }))],
  { x: 8.0, y: 4.25, w: 4.83, colW: [1.2075, 1.2075, 1.2075, 1.2075], fontSize: 13, border: { type: "solid", pt: 0.5, color: P.line }, align: "center",
    valign: "middle", rowH: 0.38, objectName: "Segment L2D table" });
card(s, 8.0, 5.2, 4.83, 1.55, P.light, "Hypotheses card");
s.addText([
  { text: "Causes to test before investing", options: { bold: true, color: P.plum, breakLine: true } },
  { text: "SDR capacity  ·  speed-to-lead  ·  lead quality  ·  demo scheduling", options: { color: P.ink } },
], { x: 8.2, y: 5.3, w: 4.45, h: 1.35, fontSize: 14, isTextBox: true, margin: 0, valign: "middle", paraSpaceAfter: 6 });
footer(s, SRC);
s.addNotes(`Leads rose ${pct(EB.gl[4], 0)} between the Feb-Jul and Aug-Oct averages; demos did not keep up. Lead-to-demo fell ${pts(EB.cd[4])} (about ${pct(Math.abs(l2dRel), 0)} relative), the largest slippage of any stage. Post-demo rates moved far less. Within segments, 1-50 and 51-100 slipped most; Enterprise held. Part of the blended drop is mix toward larger accounts. I call this a potential bottleneck, not a proven SDR capacity issue: the same pattern fits slower follow-up, weaker lead quality, or scheduling friction. Each points to a different fix, so we test first.`);

// ================================================================ 3. Finding 2
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Main" });
s.addText("51–100 delivers the highest observed ARR per lead: it converts like SMB at more than double the deal size", { placeholder: "title" });
sub(s, "12-month ARR per lead by segment (ARR ÷ high-intent leads)");
s.addChart(pres.charts.BAR, [{ name: "ARR per lead", labels: SEG.map((x, i) => `${x}\n${SEGN[i]}`), values: F.slice(0, 4).map((r) => Math.round(r.arr_lead)) }],
  { x: 0.5, y: 1.7, w: 7.0, h: 4.9, barDir: "col", chartColors: [P.soft, P.plum, P.soft, P.soft], showValue: true, dataLabelPosition: "outEnd",
    dataLabelFormatCode: "$#,##0", valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
    barGapWidthPct: 60, ...chartText, dataLabelFontSize: 13, objectName: "ARR per lead chart" });
stat(s, 8.0, 1.8, 4.8, `${pct(F[1].l2w)} vs ${pct(F[0].l2w)}`, "Lead→win for 51–100 vs 1–50: essentially the same conversion");
stat(s, 8.0, 3.15, 4.8, `${k(F[1].asp)} vs ${k(F[0].asp)}`, `ARR per logo: ${(F[1].asp / F[0].asp).toFixed(1)}× larger deals, with the same ${F[1].demos_logo.toFixed(1)} demos per logo`);
stat(s, 8.0, 4.5, 4.8, pts(EB.cd[1]), "Lead→demo change, the largest slip of any segment", P.red);
card(s, 8.0, 5.85, 4.83, 0.9, P.light, "Yield caveat");
s.addText("Revenue yield, not proven ROI: the dataset has no cost data. It is the clearest place to protect lead coverage.",
  { x: 8.15, y: 5.9, w: 4.55, h: 0.8, fontSize: 12, color: P.ink, isTextBox: true, margin: 0, valign: "middle" });
footer(s, SRC);
s.addNotes(`51-100 earns ${usd(F[1].arr_lead)} of ARR per lead versus ${usd(F[2].arr_lead)} for 101-500 and ${usd(F[3].arr_lead)} for 501+. It converts like small business but at ${(F[1].asp / F[0].asp).toFixed(1)}x the deal size. It also saw the biggest lead-to-demo slip. Careful: this is revenue yield, not ROI. We have no cost data, and larger deals may cost more to win. The point is that every lead lost here costs the most ARR.`);

// ================================================================ 4. Finding 3
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Main" });
s.addText("Growth is shifting upmarket, where conversion halves and each deal needs more support", { placeholder: "title" });
sub(s, "Change in average monthly new ARR, Feb–Jul to Aug–Oct 2021");
s.addChart(pres.charts.BAR, [{ name: "ARR growth", labels: SEG, values: EB.ga.slice(0, 4) }],
  { x: 0.5, y: 1.7, w: 5.6, h: 4.9, barDir: "col", chartColors: [P.soft, P.soft, P.plum, P.plum], showValue: true, dataLabelPosition: "outEnd",
    dataLabelFormatCode: "0%", valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
    barGapWidthPct: 60, ...chartText, dataLabelFontSize: 13, objectName: "ARR growth chart" });
const segTbl = [
  [hdr("12-month metric"), ...SEG.map(hdr)],
  ["Share of ARR", ...F.slice(0, 4).map((r) => pct(r.arr_share, 0))],
  ["Lead → win", ...F.slice(0, 4).map((r) => pct(r.l2w))],
  ["Demo → opp", ...F.slice(0, 4).map((r) => pct(r.d2o, 0))],
  ["Demos per logo", ...F.slice(0, 4).map((r) => r.demos_logo.toFixed(1))],
  ["ARR per logo", ...F.slice(0, 4).map((r) => k(r.asp))],
];
s.addTable(segTbl, { x: 6.4, y: 1.8, w: 6.43, colW: [1.83, 1.15, 1.15, 1.15, 1.15], fontSize: 13, color: P.ink, align: "center", valign: "middle",
  border: { type: "solid", pt: 0.5, color: P.line }, rowH: 0.42, objectName: "Segment metrics table" });
card(s, 6.4, 4.65, 6.43, 2.1, P.light, "So-what card");
s.addText([
  { text: "So what", options: { bold: true, color: P.plum, breakLine: true } },
  { text: `101+ employees produce ${pct(upmarketShare, 0)} of ARR and their share of leads rose from ${pct(D.mix101[0], 0)} to ${pct(D.mix101[1], 0)}.`, options: { bullet: true, breakLine: true } },
  { text: "Lower conversion and 1.5–2× the demos per logo signal complex buying: this is where SCs earn their place.", options: { bullet: true, breakLine: true } },
  { text: "1–10 is shrinking and low value: serve it with a low-touch motion, not more headcount.", options: { bullet: true } },
], { x: 6.6, y: 4.75, w: 6.1, h: 1.95, fontSize: 13, color: P.ink, isTextBox: true, margin: 0, valign: "top", paraSpaceAfter: 4 });
footer(s, SRC);
s.addNotes(`Enterprise ARR grew ${pct(EB.ga[3], 0)} and Core MM ${pct(EB.ga[2], 0)}, against single digits below 100 employees. But lead-to-win is ${pct(F[2].l2w)} in 101-500 and ${pct(F[3].l2w)} in 501+, versus about 10% below 100. Demo-to-opp falls and demos per logo rise, so these are harder, multi-stakeholder deals. That's the case for SC support upmarket, and for caution before scaling expensive Enterprise coverage. The 1-10 bucket shrank over the year; it should be low-touch.`);

// ================================================================ 5. Segmentation
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Main" });
s.addText("Four segments, run as three sales motions, with boundaries placed where the data changes", { placeholder: "title" });
const segCards = [
  ["Velocity", "1–50 employees", `${k(F[0].asp)} ARR/logo; 10% lead→win`, "High-velocity inbound; self-scheduling for 1–10", "Dedicated SDR + AE; SC by exception"],
  ["Emerging MM", "51–100 employees", `ARR/logo jumps to ${k(F[1].asp)} at the 50 line`, "Inbound with fast follow-up SLA; routed first", "Dedicated SDR + AE; pooled SC"],
  ["Core MM", "101–500 employees", `Lead→win halves to ${pct(F[2].l2w)} at the 100 line`, "Consultative; SC on about half of demos", "Dedicated SDR + AE; pooled SC"],
  ["Enterprise", "501+ employees", `Lead→win halves again; ${k(F[3].asp)} ARR/logo`, "Multi-stakeholder; SC on every deal", "Dedicated SDR + AE + SC"],
];
// motion bands
const motion = [["Velocity motion", 0.5, 2.95], ["Mid-Market motion", 3.6, 6.05], ["Enterprise motion", 9.8, 2.95]];
motion.forEach(([t, x, w]) => {
  card(s, x, 1.35, w, 0.45, P.plum, t);
  s.addText(t, { x, y: 1.35, w, h: 0.45, fontSize: 14, bold: true, color: P.white, align: "center", valign: "middle", isTextBox: true, margin: 0 });
});
segCards.forEach(([name, range, why, motionTxt, roles], i) => {
  const x = 0.5 + i * 3.1, w = 2.95;
  card(s, x, 1.95, w, 3.65, P.light, `Segment card ${name}`);
  s.addText(name, { x: x + 0.15, y: 2.05, w: w - 0.3, h: 0.4, fontSize: 18, bold: true, color: P.plum, isTextBox: true, margin: 0 });
  s.addText(range, { x: x + 0.15, y: 2.45, w: w - 0.3, h: 0.3, fontSize: 13, color: P.grey, isTextBox: true, margin: 0 });
  s.addText([
    { text: "Why here", options: { bold: true, breakLine: true } }, { text: why, options: { breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Motion", options: { bold: true, breakLine: true } }, { text: motionTxt, options: { breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Roles", options: { bold: true, breakLine: true } }, { text: roles },
  ], { x: x + 0.15, y: 2.85, w: w - 0.3, h: 2.65, fontSize: 13, color: P.ink, isTextBox: true, margin: 0, valign: "top" });
});
card(s, 0.5, 5.8, 12.33, 1.0, "FFFFFF", "Alternatives box");
s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.8, w: 12.33, h: 1.0, fill: { color: P.white }, line: { color: P.line, width: 1 }, objectName: "Alternatives outline" });
s.addText([
  { text: "Alternatives considered:  ", options: { bold: true, color: P.plum } },
  { text: `Conventional 1–100 / 101–500 / 501+ puts ${k(bk[0].asp)} and ${k(bk[2].asp)} deals in one AE queue (${optA_spread.toFixed(1)}× spread), so 51–100 competes with the smallest deals.  ` },
  { text: `1–50 / 51–500 / 501+ puts ${pct(bk[2].l2w)} and ${pct(bk[3].l2w)} lead→win in one team. Tradeoff of four segments: more routing rules and quota plans.` },
], { x: 0.7, y: 5.85, w: 11.95, h: 0.9, fontSize: 13, color: P.ink, isTextBox: true, margin: 0, valign: "middle" });
footer(s, "Boundaries can only fall on the dataset's bucket edges (10 / 50 / 100 / 500). 101–500 is one wide bucket; finer data could reveal a break near 250.");
s.addNotes("I didn't assume standard SMB/MM/ENT lines. I put a boundary wherever the data changes. At 50, deal size jumps about 2.5x while conversion stays flat. At 100, conversion halves. At 500, it halves again and deal size nearly doubles. Four segments sit inside three motions, so 51-100 and 101-500 share Mid-Market leadership and an SC pool but have separate AE teams and quotas. The conventional 1-100 split would put 51-100 deals in the same queue as $3K deals. The cost of four segments is a bit more routing and comp-plan complexity.");

// ================================================================ 6. Staffing
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Main" });
s.addText(`${hc} roles of required capacity, weighted to where the funnel needs help: first touch below 500 and SCs upmarket`, { placeholder: "title" });
sub(s, "Recommended headcount at Aug–Oct 2021 demand (average month), ROUNDUP of modeled workload");
const T = [
  [hdr("Role"), ...SEG.map((x, i) => hdr(`${SEGN[i]}\n${x}`)), hdr("Total")],
  ["SDR", ...S.h_sdr.slice(0, 4).map(String), { text: String(hcSdr), options: { bold: true } }],
  ["AE", ...S.h_ae.slice(0, 4).map(String), { text: String(hcAe), options: { bold: true } }],
  ["SC", { text: `${D.sc_pool}  (shared pool)`, options: { colspan: 3 } }, String(D.sc_ent), { text: String(hcSc), options: { bold: true } }],
  [{ text: "Total", options: { bold: true } }, { text: "", options: { colspan: 4 } }, { text: String(hc), options: { bold: true, color: P.plum } }],
  ["AEs per SDR", ...S.ae_per_sdr.slice(0, 4).map((v) => v.toFixed(1)), S.ae_per_sdr[4].toFixed(1)],
  ["AEs per SC", { text: D.ae_per_sc_pool.toFixed(1), options: { colspan: 3 } }, D.ae_per_sc_ent.toFixed(1), (hcAe / hcSc).toFixed(1)],
  ["Run-rate ARR", ...S.arr_y.slice(0, 4).map((v) => m(v)), { text: m(base12), options: { bold: true } }],
];
s.addTable(T, { x: 0.5, y: 1.75, w: 8.4, colW: [1.6, 1.36, 1.36, 1.36, 1.36, 1.36], fontSize: 13, color: P.ink, align: "center", valign: "middle",
  border: { type: "solid", pt: 0.5, color: P.line }, rowH: [0.6, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45], objectName: "Staffing table" });
s.addText("Why: Velocity and 51–100 get dedicated SDR coverage where lead→demo slipped most; Core MM gets the largest AE team; Enterprise gets an SC on every deal.",
  { x: 0.5, y: 5.6, w: 8.4, h: 0.75, fontSize: 13, color: P.ink, isTextBox: true, margin: 0, valign: "top" });
card(s, 9.2, 1.75, 3.63, 2.6, P.light, "Method card");
s.addText([
  { text: "How it is sized", options: { bold: true, color: P.plum, breakLine: true } },
  { text: "SDR = leads ÷ (capacity × 85%)", options: { breakLine: true } },
  { text: "AE = demos ÷ (capacity × 85%)", options: { breakLine: true } },
  { text: "SC = demos × SC rate ÷ (capacity × 85%)", options: { breakLine: true } },
  { text: "Round up; SC pool rounds up the combined need", options: { color: P.grey } },
], { x: 9.35, y: 1.85, w: 3.35, h: 2.4, fontSize: 13, color: P.ink, isTextBox: true, margin: 0, valign: "top", paraSpaceAfter: 4 });
card(s, 9.2, 4.5, 3.63, 1.6, P.plum, "Upside card");
s.addText([
  { text: "Upside sensitivity", options: { bold: true, breakLine: true } },
  { text: `If lead→demo recovers to Feb–Jul levels: +${m(up12 - base12)} ARR and +${D.d_hc_tot} AE (Core MM), ${D.u_hc_tot} roles. A scenario, not a forecast.` },
], { x: 9.35, y: 4.6, w: 3.35, h: 1.4, fontSize: 13, color: P.white, isTextBox: true, margin: 0, valign: "top", paraSpaceAfter: 4 });
footer(s, SRC_STAFF);
s.addNotes(`This is required capacity at Aug-Oct demand, not a hiring plan. We don't know current headcount, so the action is to compare this to the current team and hire only the gap. The math is one line per role, using illustrative capacity assumptions and 85% utilization, rounded up so the team covers modeled workload. Below 100 employees, SC needs are fractional, so one shared SC pool covers Velocity and Mid-Market. Enterprise gets dedicated SCs at roughly ${D.ae_per_sc_ent.toFixed(1)} AEs per SC. If lead-to-demo recovers, demos rise about ${n0(S.u_demos[4] - S.demos[4])} a month, adding one Core MM AE and about ${m(up12 - base12)} of ARR. That upside depends on the cause of the decline. I'm not claiming this mix is optimal; it's what these assumptions imply, and the model lets you change them.`);

// ================================================================ 7. Recommendations & 30/60/90
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Main" });
s.addText("Recommendations: diagnose first touch, organize by segment, and test before scaling headcount", { placeholder: "title" });
const recs = [
  ["Fix lead→demo before buying more pipeline", "Instrument speed-to-lead and test SLAs, routing and self-scheduling, starting with 1–50 and 51–100."],
  ["Organize into four segments and three motions", "Dedicated AE teams by segment; route 51–100 leads first; one shared SC pool below 500 employees."],
  [`Staff to ${hc} roles of required capacity`, "Compare with current headcount and hire only the gap. Add the upside AE only if lead→demo recovers."],
  ["Invest upmarket with discipline", "SC on every Enterprise deal; hold Enterprise expansion until win rates improve. No new headcount for 1–10."],
];
recs.forEach(([h, b], i) => {
  const y = 1.45 + i * 1.33;
  badge(s, i + 1, 0.5, y + 0.05, 0.45);
  s.addText(h, { x: 1.1, y, w: 5.6, h: 0.4, fontSize: 15, bold: true, color: P.plum, isTextBox: true, margin: 0 });
  s.addText(b, { x: 1.1, y: y + 0.4, w: 5.6, h: 0.8, fontSize: 13, color: P.ink, isTextBox: true, margin: 0, valign: "top" });
});
const plan = [
  ["30 days: diagnose", "RevOps + Marketing Ops", "Instrument speed-to-lead and lead→demo by segment and source; audit a sample of leads that never booked; launch four-segment routing.",
   "KPI: first-touch time; weekly lead→demo by segment"],
  ["60 days: pilot", "Sales + SDR leadership", "Run a speed-to-lead SLA and self-scheduling pilot vs control in 1–50 and 51–100; set SC engagement rules for Core MM and Enterprise.",
   "KPI: lead→demo vs Feb–Jul level; demo show rate"],
  ["90 days: decide", "CRO + Finance + RevOps", "Compare current headcount to required capacity; hire the gap where the pilot shows lift; set quotas by segment; review Enterprise win rate.",
   "KPI: lead→demo, opp→win, ARR per AE by segment"],
];
plan.forEach(([h, own, what, kpi], i) => {
  const y = 1.45 + i * 1.78;
  card(s, 7.0, y, 5.83, 1.62, P.light, `Plan ${h}`);
  s.addText([{ text: h, options: { bold: true, color: P.plum } }, { text: `   Owner: ${own}`, options: { color: P.grey } }],
    { x: 7.15, y: y + 0.08, w: 5.55, h: 0.35, fontSize: 14, isTextBox: true, margin: 0 });
  s.addText(what, { x: 7.15, y: y + 0.45, w: 5.55, h: 0.75, fontSize: 12, color: P.ink, isTextBox: true, margin: 0, valign: "top" });
  s.addText(kpi, { x: 7.15, y: y + 1.2, w: 5.55, h: 0.35, fontSize: 12, italic: true, color: P.teal, isTextBox: true, margin: 0 });
});
footer(s, "Owners and KPIs are proposals for discussion. Targets should be set once the 30-day diagnostic establishes current baselines.");
s.addNotes("Four recommendations. One: before adding pipeline spend or SDRs, find out why lead-to-demo slipped; the 30-day diagnostic separates capacity from lead quality. Two: organize by the four segments, with 51-100 routed first. Three: staff to required capacity and hire only the gap; add the upside AE only if the pilot shows recovery. Four: invest upmarket with discipline. Enterprise gets full SC support but no expansion until win rates move, and 1-10 goes low-touch. Where I'd avoid overinvesting: Enterprise AE expansion, 1-10 headcount, and more lead generation before the first-touch issue is understood.");

// ================================================================ Appendix
pres.addSection({ title: "Appendix" });
s = pres.addSlide({ masterName: "DARK", sectionTitle: "Appendix" });
s.addText("Appendix", { placeholder: "title" });
s.addText("A1 Assumptions and formulas  ·  A2 Staffing calculation  ·  A3 Funnel by bucket  ·  A4 Segmentation options  ·  A5 Illustrative cost  ·  A6 Data notes", { placeholder: "body" });

// A1
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Appendix" });
s.addText("A1. Five staffing assumptions drive the model; all are illustrative and editable", { placeholder: "title" });
const A1 = [
  [hdr("Assumption (per ramped person per month)"), ...SEG.map((x, i) => hdr(`${SEGN[i]} ${x}`)), hdr("Rationale")],
  ["SDR lead capacity", ...S.sdr_cap.slice(0, 4).map(n0), "Bigger accounts need more research per lead"],
  ["AE demo capacity", ...S.ae_cap.slice(0, 4).map(n0), "Enterprise demos are multi-session"],
  ["SC engagement rate", ...S.sc_rate.slice(0, 4).map((v) => pct(v, 0)), "Rises with complexity in the data"],
  ["SC engagement capacity", ...S.sc_cap.slice(0, 4).map(n0), "Enterprise engagements are deeper"],
  ["Target utilization", ...S.util.slice(0, 4).map((v) => pct(v, 0)), "Buffer for PTO, training, spikes"],
];
s.addTable(A1, { x: 0.5, y: 1.4, w: 12.33, colW: [3.0, 1.55, 1.55, 1.55, 1.55, 3.13], fontSize: 13, color: P.ink, align: "center", valign: "middle",
  border: { type: "solid", pt: 0.5, color: P.line }, rowH: 0.48, objectName: "Assumptions table" });
card(s, 0.5, 4.55, 12.33, 1.75, P.light, "Formula card");
s.addText([
  { text: "Formulas (Staffing tab)", options: { bold: true, color: P.plum, breakLine: true } },
  { text: "Demand = average month of Aug–Oct 2021 by segment (Nov–Dec dip excluded; seen once, so not confirmed seasonality)", options: { bullet: true, breakLine: true } },
  { text: "SDR FTE = leads ÷ (SDR capacity × utilization);  AE FTE = demos ÷ (AE capacity × utilization);  SC FTE = demos × SC rate ÷ (SC capacity × utilization)", options: { bullet: true, breakLine: true } },
  { text: "Headcount = ROUNDUP per segment; shared SC pool (1–500) = ROUNDUP of combined FTE;  ARR = demos × demo→win × ARR per logo × 12", options: { bullet: true, breakLine: true } },
  { text: `Most sensitive: AE demo capacity. At ±20% capacity, AE headcount ranges ${aeAt(1.2)}–${aeAt(0.8)} (base ${hcAe}).`, options: { bullet: true } },
], { x: 0.7, y: 4.65, w: 11.95, h: 1.55, fontSize: 13, color: P.ink, isTextBox: true, margin: 0, valign: "top", paraSpaceAfter: 4 });
footer(s, SRC_STAFF);

// A2
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Appendix" });
s.addText("A2. Staffing calculation: demand → fractional FTE → recommended headcount, baseline and upside", { placeholder: "title" });
const f2 = (v) => v.toFixed(2);
const A2 = [
  [hdr("Line"), ...SEG.map(hdr), hdr("Total")],
  ["Leads / month", ...S.leads.map(n0)],
  ["Demos / month", ...S.demos.map(n0)],
  ["Lead → demo", ...S.l2d.map((v) => pct(v))],
  ["SDR FTE → headcount", ...S.f_sdr.slice(0, 4).map((v, i) => `${f2(v)} → ${S.h_sdr[i]}`), `${f2(S.f_sdr[4])} → ${hcSdr}`],
  ["AE FTE → headcount", ...S.f_ae.slice(0, 4).map((v, i) => `${f2(v)} → ${S.h_ae[i]}`), `${f2(S.f_ae[4])} → ${hcAe}`],
  ["SC FTE → headcount", { text: `${f2(S.f_sc[0] + S.f_sc[1] + S.f_sc[2])} → ${D.sc_pool} (pool)`, options: { colspan: 3 } }, `${f2(S.f_sc[3])} → ${D.sc_ent}`, `${f2(S.f_sc[4])} → ${hcSc}`],
  ["AE utilization at headcount", ...S.ae_util.map((v) => pct(v, 0))],
  ["Upside lead → demo", ...S.u_l2d.map((v) => pct(v))],
  ["Upside AE FTE → headcount", ...S.u_f_ae.slice(0, 4).map((v, i) => `${f2(v)} → ${S.u_h_ae[i]}`), `${f2(S.u_f_ae[4])} → ${S.u_h_ae[4]}`],
  ["Upside SC headcount", { text: `${D.u_sc_pool} (pool)`, options: { colspan: 3 } }, String(D.u_sc_ent), String(D.u_sc_tot)],
  ["ARR: baseline → upside", ...S.arr_y.map((v, i) => `${m(v)} → ${m(S.u_arr_y[i])}`)],
];
s.addTable(A2, { x: 0.5, y: 1.35, w: 12.33, colW: [2.73, 1.92, 1.92, 1.92, 1.92, 1.92], fontSize: 12, color: P.ink, align: "center", valign: "middle",
  border: { type: "solid", pt: 0.5, color: P.line }, rowH: 0.42, objectName: "Calculation table" });
footer(s, `Illustrative assumptions. Upside total: ${D.u_hc_tot} roles; SDRs unchanged because they are sized to leads.`);

// A3
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Appendix" });
s.addText("A3. Twelve-month funnel by employee bucket: conversion breaks at 100 and 500, deal size at 50", { placeholder: "title" });
const A3 = [[hdr("Bucket"), hdr("Leads"), hdr("Demos"), hdr("Opps"), hdr("Logos"), hdr("ARR"), hdr("Lead→\ndemo"), hdr("Demo→\nopp"), hdr("Opp→\nwin"),
  hdr("Lead→\nwin"), hdr("ARR/logo"), hdr("ARR/lead"), hdr("ARR share")]];
[...D.bucket_names, "Total"].forEach((b, i) => {
  const r = bk[i], bold = b === "Total";
  A3.push([b, n0(r.leads), n0(r.demos), n0(r.qo), n0(r.logos), m(r.arr), pct(r.l2d), pct(r.d2o), pct(r.o2w), pct(r.l2w), usd(r.asp), usd(r.arr_lead),
    pct(r.arr_share)].map((t) => ({ text: t, options: { bold } })));
});
s.addTable(A3, { x: 0.5, y: 1.4, w: 12.33, colW: [0.95, 0.85, 0.85, 0.85, 0.8, 0.95, 1.05, 1.05, 0.95, 0.95, 1.0, 1.0, 1.08], fontSize: 12, color: P.ink,
  align: "center", valign: "middle", border: { type: "solid", pt: 0.5, color: P.line }, rowH: 0.42, objectName: "Bucket funnel table" });
card(s, 0.5, 4.75, 12.33, 1.5, P.light, "Bucket notes");
const b10 = D.monthly; // monthly totals for headline growth
s.addText([
  { text: `1–10 and 11–50 convert identically (${pct(bk[0].l2w)} lead→win); 1–10 earns ${usd(bk[0].asp)} per logo and its leads shrank over the year.`, options: { bullet: true, breakLine: true } },
  { text: `Total 12-month new ARR ${m(bk[5].arr)}; monthly new ARR rose from ${m(b10[0].arr)} (Feb-21) to a record ${m(b10[11].arr)} (Jan-22).`, options: { bullet: true, breakLine: true } },
  { text: "Nov–Dec 2021: leads fell ~47% vs Oct and all rates dipped. Excluded from the staffing baseline; if it recurs, annual ARR is below the ×12 run-rate.", options: { bullet: true } },
], { x: 0.7, y: 4.85, w: 11.95, h: 1.3, fontSize: 13, color: P.ink, isTextBox: true, margin: 0, valign: "top", paraSpaceAfter: 4 });
footer(s, SRC);

// A4
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Appendix" });
s.addText("A4. Segmentation options compared: the four-segment model keeps each team within one deal size and conversion profile", { placeholder: "title" });
const A4 = [
  [hdr("Option"), hdr("Segments"), hdr("Strength"), hdr("Weakness")],
  ["Recommended: 4 segments / 3 motions", "1–50 · 51–100 · 101–500 · 501+", "Each AE team sits in one conversion regime and deal-size band; protects 51–100",
   "More routing rules and quota plans"],
  ["A. Conventional 3-tier", "1–100 · 101–500 · 501+", "Simplest; uniform conversion within tiers",
   `Mixes ${usd(bk[0].asp)} and ${usd(bk[2].asp)} ARR/logo (${optA_spread.toFixed(1)}× spread) in one AE queue`],
  ["B. Alternative 3-tier", "1–50 · 51–500 · 501+", "Splits on deal size",
   `Mixes ${pct(bk[2].l2w)} and ${pct(bk[3].l2w)} lead→win (${optB_spread.toFixed(1)}×) in one Mid-Market team`],
];
s.addTable(A4, { x: 0.5, y: 1.5, w: 12.33, colW: [3.0, 2.6, 3.4, 3.33], fontSize: 13, color: P.ink, valign: "middle",
  border: { type: "solid", pt: 0.5, color: P.line }, rowH: [0.45, 0.85, 0.85, 0.85], objectName: "Options table" });
footer(s, "Boundaries can only fall on dataset bucket edges. Finer employee-count data (especially inside 101–500) would let us test other cut points.");

// A5
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Appendix" });
s.addText("A5. Optional, illustrative: required capacity implies roughly " + m(D.cost.base[3], 1) + " of annual fully loaded cost", { placeholder: "title" });
const roles = ["SDR", "AE (blended)", "SC"], hcs = [hcSdr, hcAe, hcSc], uhcs = [S.u_h_sdr[4], S.u_h_ae[4], D.u_sc_tot];
const A5 = [[hdr("Role"), hdr("Illustrative loaded cost / FTE"), hdr("Baseline headcount"), hdr("Baseline cost"), hdr("Upside headcount"), hdr("Upside cost")]];
roles.forEach((r, i) => A5.push([r, usd(D.cost.unit[i]), String(hcs[i]), m(D.cost.base[i]), String(uhcs[i]), m(D.cost.up[i])]));
A5.push(["Total", "", String(hc), m(D.cost.base[3]), String(D.u_hc_tot), m(D.cost.up[3])].map((t) => ({ text: t, options: { bold: true } })));
s.addTable(A5, { x: 0.5, y: 1.5, w: 12.33, colW: [2.2, 2.2, 1.95, 2.0, 1.95, 2.03], fontSize: 13, color: P.ink, align: "center", valign: "middle",
  border: { type: "solid", pt: 0.5, color: P.line }, rowH: 0.48, objectName: "Cost table" });
card(s, 0.5, 4.4, 12.33, 1.1, P.light, "Cost caveat");
s.addText([
  { text: "Scale only. San Francisco loaded cost (salary + target variable + benefits) is illustrative and should be validated with Comp.", options: { bullet: true, breakLine: true } },
  { text: "Current headcount is unknown, so this is not incremental spend. It is not an ROI, CAC or payback analysis; the dataset has no cost data.", options: { bullet: true } },
], { x: 0.7, y: 4.5, w: 11.95, h: 0.9, fontSize: 13, color: P.ink, isTextBox: true, margin: 0, valign: "top", paraSpaceAfter: 4 });
footer(s, SRC_STAFF);

// A6
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Appendix" });
s.addText("A6. Data notes, limitations and what we would want next", { placeholder: "title" });
const cols = [
  ["Data checks", ["60 monthly rows: 12 months × 5 buckets, no gaps", "No missing values or duplicates", "Funnel ordered in every row (leads ≥ demos ≥ opps ≥ logos)",
    "All monthly; no annual records to reconcile"]],
  ["Limitations", ["No headcount, so required capacity ≠ hires", "Same-month ratios, not cohort conversion", "One year: Nov–Dec dip not confirmed as seasonal",
    "101–500 is one wide bucket", "No cost data: ARR per lead is not ROI"]],
  ["Data we would want", ["Current SDR / AE / SC headcount and ramp status", "Speed-to-lead and lead source by segment", "Closed-lost reasons and sales-cycle length",
    "Rep-level activity and capacity", "Retention / NRR by segment"]],
];
cols.forEach(([h, items], i) => {
  const x = 0.5 + i * 4.18, w = 3.97;
  card(s, x, 1.45, w, 3.9, P.light, `Notes ${h}`);
  s.addText(h, { x: x + 0.2, y: 1.6, w: w - 0.4, h: 0.4, fontSize: 17, bold: true, color: P.plum, isTextBox: true, margin: 0 });
  s.addText(items.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < items.length - 1 } })),
    { x: x + 0.2, y: 2.1, w: w - 0.4, h: 3.1, fontSize: 14, color: P.ink, isTextBox: true, margin: 0, valign: "top", paraSpaceAfter: 8 });
});
footer(s, SRC);

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
