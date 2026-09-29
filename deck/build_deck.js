// Builds deck/portfolio_recommendation.pptx
const pptxgen = require("pptxgenjs");

const OUT = process.argv[2] || "portfolio_recommendation.pptx";

// Palette
const NAVY = "0F1033", INDIGO = "4A4AF4", LAV = "ECECFE", LAV2 = "D9D9FC",
      INK = "1A1A2E", MUTED = "5B5B73", LINE = "DADAF0", WHITE = "FFFFFF",
      RED = "C93B3B", GREEN = "1E8A4C";
const FONT = "Arial";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Where to invest next year";

const W = 13.333, M = 0.6;

function title(slide, text, opts = {}) {
  slide.addText(text, { x: M, y: 0.45, w: W - 2 * M, h: 0.95, fontFace: FONT, fontSize: 26, bold: true,
    color: opts.color || INK, valign: "top", margin: 0, isTextBox: true });
}
function kicker(slide, text, color = INDIGO) {
  slide.addText(text.toUpperCase(), { x: M, y: 0.18, w: 8, h: 0.3, fontFace: FONT, fontSize: 11, bold: true,
    color, charSpacing: 2, margin: 0, isTextBox: true });
}
function footnote(slide, text, color = MUTED) {
  slide.addText(text, { x: M, y: 6.95, w: W - 2 * M, h: 0.35, fontFace: FONT, fontSize: 10, color,
    margin: 0, valign: "bottom", isTextBox: true });
}
function card(slide, x, y, w, h, fill = LAV) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.12 });
}
function stat(slide, x, y, w, big, label, bigColor = INDIGO, bigSize = 30) {
  slide.addText(big, { x, y, w, h: 0.6, fontFace: FONT, fontSize: bigSize, bold: true, color: bigColor, margin: 0, isTextBox: true });
  slide.addText(label, { x, y: y + 0.62, w, h: 0.62, fontFace: FONT, fontSize: 12, color: INK, margin: 0, valign: "top", isTextBox: true });
}
function pill(slide, x, y, text, fill, color = WHITE, w = 1.3) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.3, fill: { color: fill }, line: { color: fill }, rectRadius: 0.15 });
  slide.addText(text, { x, y, w, h: 0.3, fontFace: FONT, fontSize: 10, bold: true, color, align: "center", valign: "middle", margin: 0, isTextBox: true });
}

// ---------- 1. Title ----------
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 8.9, y: -1.6, w: 6.2, h: 6.2, fill: { color: INDIGO }, line: { color: INDIGO } });
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: 3.6, w: 3.4, h: 3.4, fill: { color: "2B2CA8" }, line: { color: "2B2CA8" } });
  s.addText("RISK & ANALYTICS  ·  PORTFOLIO REVIEW", { x: M, y: 1.7, w: 8, h: 0.35, fontFace: FONT, fontSize: 12, bold: true,
    color: LAV2, charSpacing: 2, margin: 0, isTextBox: true });
  s.addText("Where to invest next year", { x: M, y: 2.15, w: 8.3, h: 1.1, fontFace: FONT, fontSize: 44, bold: true,
    color: WHITE, margin: 0, isTextBox: true });
  s.addText("Grow low-loss verticals and repeat customers; tighten the riskiest travel credit",
    { x: M, y: 3.3, w: 7.8, h: 0.9, fontFace: FONT, fontSize: 20, color: LAV2, margin: 0, valign: "top", isTextBox: true });
  const chips = [["$4.0M", "GMV"], ["7,134", "loans"], ["4,000", "consumers"], ["7.0%", "contribution margin"]];
  chips.forEach(([big, lab], i) => {
    const x = M + i * 1.95;
    s.addText(big, { x, y: 5.0, w: 1.85, h: 0.5, fontFace: FONT, fontSize: 24, bold: true, color: WHITE, margin: 0, isTextBox: true });
    s.addText(lab, { x, y: 5.5, w: 1.85, h: 0.35, fontFace: FONT, fontSize: 12, color: LAV2, margin: 0, isTextBox: true });
  });
  footnote(s, "Loans originated Jul 2024 – Dec 2025, snapshot 31 Dec 2025", LAV2);
  s.addNotes("Goal: where to add and where to pull back capital next year. One-line answer: grow healthcare and home furnishing merchants, invest in turning first-time customers into repeat customers, and tighten credit for the riskiest travel borrowers. The portfolio in scope: $4.0M GMV, 7,134 loans, 4,000 consumers, ~7% contribution margin after losses and funding.");
}

// ---------- 2. Executive summary ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Recommendation");
  title(s, "Shift growth to where losses are low and customers come back");

  const items = [
    ["INVEST", GREEN, "Merchant acquisition: healthcare & home furnishing",
      "Best risk-adjusted margins (9.1%, 7.4%) with low loss. One new merchant ≈ +25–35% vertical GMV."],
    ["INVEST", GREEN, "Repeat customers",
      "Repeat loans lose 2.5× less; 82% of loss comes from first loans. Lift big-ticket repeat rate and steer second loans to interest-bearing."],
    ["TIGHTEN", RED, "Travel borrowers with FICO < 640",
      "~40% delinquent or charged off, 6.3% loss. Raise the risk model cut-off and lower limits: −25% to −50% volume."],
    ["REPRICE", INDIGO, "0% APR merchant fee",
      "Fee (5.0%) is below interest-bearing (5.4%) at the same merchants. Raise to at least parity."],
  ];
  items.forEach(([tag, col, head, body], i) => {
    const y = 1.55 + i * 1.3;
    card(s, M, y, 7.6, 1.15, "F6F6FE");
    pill(s, M + 0.25, y + 0.2, tag, col);
    s.addText(head, { x: M + 1.75, y: y + 0.13, w: 5.7, h: 0.4, fontFace: FONT, fontSize: 15, bold: true, color: INK, margin: 0, isTextBox: true });
    s.addText(body, { x: M + 1.75, y: y + 0.52, w: 5.7, h: 0.58, fontFace: FONT, fontSize: 12, color: MUTED, margin: 0, valign: "top", isTextBox: true });
  });

  const px = 8.65, pw = W - M - px;
  card(s, px, 1.55, pw, 5.05, NAVY);
  s.addText("One-year impact", { x: px + 0.35, y: 1.8, w: pw - 0.7, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: LAV2, margin: 0, isTextBox: true });
  s.addText("conservative → ambitious", { x: px + 0.35, y: 2.15, w: pw - 0.7, h: 0.3, fontFace: FONT, fontSize: 11, color: LAV2, margin: 0, isTextBox: true });
  const stats = [["+7% → +15%", "profit (contribution)"], ["+4% → +8%", "GMV"], ["2.32% → 2.1–2.2%", "credit loss rate"]];
  stats.forEach(([big, lab], i) => {
    const y = 2.7 + i * 1.25;
    s.addText(big, { x: px + 0.35, y, w: pw - 0.7, h: 0.55, fontFace: FONT, fontSize: 26, bold: true, color: WHITE, margin: 0, isTextBox: true });
    s.addText(lab, { x: px + 0.35, y: y + 0.55, w: pw - 0.7, h: 0.35, fontFace: FONT, fontSize: 12, color: LAV2, margin: 0, isTextBox: true });
  });
  footnote(s, "Impact sized on 2025 volume at seasoned (Jan 2024–Jun 2025) margins and loss rates; 0% APR repricing not included in totals.");
  s.addNotes("Four moves. Two where we invest: merchant acquisition in healthcare and home furnishing, and converting first-time customers into repeat customers. One where we pull back: travel borrowers below 640 FICO, via the risk model rather than a blunt cut-off. One pricing fix: 0% APR merchant fee. Together: profit grows faster than GMV and the loss rate goes down, because growth is steered toward low-loss segments.");
}

// ---------- 3. Portfolio health ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Portfolio health");
  title(s, "The book grew 33% year over year with stable credit quality and a ~7% margin");

  s.addChart(pres.charts.BAR, [{ name: "GMV ($K)", labels: ["Q3-24", "Q4-24", "Q1-25", "Q2-25", "Q3-25", "Q4-25"],
    values: [459, 664, 650, 704, 721, 771] }], {
    x: M, y: 1.5, w: 7.2, h: 4.2, barDir: "col", chartColors: [INDIGO],
    showTitle: true, title: "GMV by origination quarter ($K)", titleFontFace: FONT, titleFontSize: 13, titleColor: INK,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelColor: INK, dataLabelFontFace: FONT,
    catAxisLabelColor: MUTED, valAxisLabelColor: MUTED, catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
    catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT,
    valGridLine: { color: "E6E6EE", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 60,
  });

  const px = 8.3, pw = W - M - px;
  s.addText("Unit economics, % of GMV", { x: px, y: 1.55, w: pw, h: 0.35, fontFace: FONT, fontSize: 13, bold: true, color: INK, margin: 0, isTextBox: true });
  const rows = [["Revenue (merchant fee + interest)", "10.0%", INK], ["− Credit loss", "2.0%", RED], ["− Funding cost", "1.1%", RED]];
  rows.forEach(([lab, val, col], i) => {
    const y = 2.0 + i * 0.55;
    s.addText(lab, { x: px, y, w: pw - 1.2, h: 0.45, fontFace: FONT, fontSize: 13, color: INK, margin: 0, valign: "middle", isTextBox: true });
    s.addText(val, { x: px + pw - 1.2, y, w: 1.2, h: 0.45, fontFace: FONT, fontSize: 16, bold: true, color: col, align: "right", margin: 0, valign: "middle", isTextBox: true });
  });
  card(s, px, 3.72, pw, 0.6, LAV);
  s.addText("= Contribution margin", { x: px + 0.15, y: 3.72, w: pw - 1.4, h: 0.6, fontFace: FONT, fontSize: 13, bold: true, color: INK, margin: 0, valign: "middle", isTextBox: true });
  s.addText("7.0%", { x: px + pw - 1.35, y: 3.72, w: 1.2, h: 0.6, fontFace: FONT, fontSize: 20, bold: true, color: INDIGO, align: "right", margin: 0, valign: "middle", isTextBox: true });

  s.addText([
    { text: "Growth mix: interest-bearing +64% (highest margin), fashion +84% (lowest margin)", options: { bullet: true, breakLine: true } },
    { text: "Average FICO flat at ~652 — growth didn't come from loosening credit", options: { bullet: true, breakLine: true } },
    { text: "Holiday vintage (Q4-24): 10% of loans went bad, 2× other quarters", options: { bullet: true } },
  ], { x: px, y: 4.55, w: pw, h: 2.0, fontFace: FONT, fontSize: 12, color: INK, paraSpaceAfter: 8, margin: 0, valign: "top", isTextBox: true });
  footnote(s, "YoY = Jul–Dec 2025 vs Jul–Dec 2024. Funding cost = cost of funds × term × average outstanding balance (½ principal).");
  s.addNotes("Healthy trajectory: GMV up every quarter, 33% year over year for the second half. Credit quality is stable: average FICO ~652 throughout. Economics: revenue 10% of GMV, loss 2%, funding 1%, leaving ~7% contribution. Two things to flag: growth is led by interest-bearing, which is good, but also by fashion, our lowest-margin vertical; and the Q4-2024 holiday vintage had double the bad rate. Recent quarters show lower margins only because those loans are young.");
}

// ---------- 4. Risk segmentation ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Where value and risk sit");
  title(s, "Travel carries the risk, fashion dilutes the margin, healthcare and home furnishing lead");

  const hdr = ["Vertical", "Share of GMV", "Bad rate", "Loss rate", "Margin", "Share of profit"];
  const data = [
    ["Healthcare", "15%", "4%", "1.5%", "9.1%", "20%"],
    ["Travel", "20%", "21%", "3.5%", "8.7%", "24%"],
    ["Fitness", "7%", "5%", "1.2%", "7.4%", "7%"],
    ["Home furnishing", "14%", "4%", "1.2%", "7.4%", "15%"],
    ["Electronics", "20%", "4%", "1.4%", "6.4%", "18%"],
    ["Auto parts", "7%", "4%", "1.9%", "6.1%", "6%"],
    ["Fashion", "18%", "5%", "2.5%", "3.8%", "10%"],
  ];
  const highlight = { "Travel": { 2: RED, 3: RED }, "Fashion": { 3: RED, 4: RED }, "Healthcare": { 4: GREEN }, "Home furnishing": { 3: GREEN, 4: GREEN } };
  const rows = [hdr.map((h, j) => ({ text: h, options: { bold: true, color: WHITE, fill: { color: NAVY }, align: j ? "center" : "left" } }))];
  data.forEach((r, i) => rows.push(r.map((c, j) => {
    const col = (highlight[r[0]] || {})[j];
    return { text: c, options: { color: col || INK, bold: !!col || j === 0, align: j ? "center" : "left", fill: { color: i % 2 ? "F6F6FE" : WHITE } } };
  })));
  s.addTable(rows, { x: M, y: 1.6, w: 7.9, colW: [1.9, 1.2, 1.1, 1.1, 1.1, 1.5], rowH: 0.52, fontFace: FONT, fontSize: 13,
    border: { type: "solid", pt: 0.5, color: LINE }, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1] });

  const px = 8.9, pw = W - M - px;
  const callouts = [
    ["Travel, FICO < 640", "~40% delinquent or charged off; 6.3% loss. Margin 6.7% falls to ~0% if delinquent loans charge off.", RED],
    ["Fashion", "51% of loans but 10% of profit. Small tickets; fashion Split Pay earns ~1% margin.", RED],
    ["By product", "Interest-bearing is 58% of GMV and 82% of profit. Split Pay (2.6%) and 0% APR (3.4%) margins are thin.", INDIGO],
  ];
  callouts.forEach(([h, b, col], i) => {
    const y = 1.6 + i * 1.72;
    card(s, px, y, pw, 1.55, "F6F6FE");
    s.addText(h, { x: px + 0.2, y: y + 0.12, w: pw - 0.4, h: 0.38, fontFace: FONT, fontSize: 14, bold: true, color: col, margin: 0, isTextBox: true });
    s.addText(b, { x: px + 0.2, y: y + 0.52, w: pw - 0.4, h: 0.95, fontFace: FONT, fontSize: 12, color: INK, margin: 0, valign: "top", isTextBox: true });
  });
  footnote(s, "Bad rate = share of loans delinquent 30/60 or charged off. Rates are % of GMV. Ranking holds on seasoned loans (originated by Jun 2025).");
  s.addNotes("Sorted by margin. Travel has five times the bad rate of any other vertical, but it is still our largest profit pool because revenue is high. The risk is concentrated in sub-640 FICO travel borrowers: 40% bad rate, and the margin disappears if the delinquent loans charge off. Fashion is half our loans but a tenth of profit. Healthcare and home furnishing combine low loss with high revenue. By product, interest-bearing earns 82% of profit.");
}

// ---------- 5. Repeat customers ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Repeat customers");
  title(s, "Repeat customers are the cheapest growth: 2.5× lower loss, and most move up to interest-bearing");

  s.addChart(pres.charts.BAR, [{ name: "Repeat within 180 days", labels: ["Healthcare", "Auto parts", "Travel", "Home furnishing", "Fashion", "Electronics", "Fitness"],
    values: [17, 17, 20, 23, 94, 95, 96] }], {
    x: M, y: 1.55, w: 6.6, h: 4.9, barDir: "bar", chartColors: [INDIGO],
    showTitle: true, title: "Customers who take a 2nd loan within 180 days (%), by first-loan vertical", titleFontFace: FONT, titleFontSize: 12, titleColor: INK,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelColor: INK, dataLabelFontFace: FONT,
    catAxisLabelColor: INK, valAxisLabelColor: MUTED, catAxisLabelFontSize: 12, valAxisLabelFontSize: 10,
    catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, valAxisMaxVal: 100,
    valGridLine: { color: "E6E6EE", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 50,
  });

  const px = 7.7, pw = W - M - px;
  const stats = [
    ["2.6% vs 1.0%", "Loss rate, first loans vs repeat loans — lower in every vertical and product"],
    ["82%", "of credit loss dollars come from first loans"],
    ["63%", "of customers who start with Split Pay take an interest-bearing loan next"],
  ];
  stats.forEach(([big, lab], i) => {
    const y = 1.55 + i * 1.45;
    card(s, px, y, pw, 1.3, LAV);
    s.addText(big, { x: px + 0.25, y: y + 0.12, w: pw - 0.5, h: 0.55, fontFace: FONT, fontSize: 26, bold: true, color: INDIGO, margin: 0, isTextBox: true });
    s.addText(lab, { x: px + 0.25, y: y + 0.66, w: pw - 0.5, h: 0.58, fontFace: FONT, fontSize: 12, color: INK, margin: 0, valign: "top", isTextBox: true });
  });
  s.addText("Repeat behaviour is set by the vertical, not the product. Big-ticket verticals have the most room to grow it.",
    { x: px, y: 5.95, w: pw, h: 0.6, fontFace: FONT, fontSize: 12, italic: true, color: MUTED, margin: 0, valign: "top", isTextBox: true });
  footnote(s, "Customers whose first loan was by Jun 2025. Synthetic data: repeat timing is very regular (100–185 days), so read sizes as indicative.");
  s.addNotes("A customer who repays their first loan is a much better risk the second time: loss falls from 2.6% to 1.0%, and that holds in every vertical and product, including on seasoned loans. 82% of loss dollars come from first loans. Repeat rates are set by the vertical: frequent-purchase verticals like fashion and electronics already get ~95%; big-ticket verticals only ~20%, which is where the upside is. And Split Pay works as an entry product: most customers move to interest-bearing, our highest-margin product, on their next loan.");
}

// ---------- 6. Recommendations & sizing ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Sizing");
  title(s, "Together the moves add ~7–15% profit while lowering the loss rate");

  const H = (t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, align: "center" } });
  const rows = [
    [H("Recommendation"), H("Assumption (conservative / ambitious)"), H("GMV"), H("Profit"), H("Credit loss $")],
  ];
  const body = [
    ["1. Grow healthcare & home furnishing", "+10% / +20% vertical GMV (≈ 1 new merchant each at +20%)", "+$78K / +$155K", "+$7.0K / +$14.0K", "+$1.3K / +$2.7K"],
    ["2. Repeat customers", "+5 / +10 pts big-ticket repeat rate; 5% / 10% of repeat loans move Split Pay → interest-bearing", "+$86K / +$172K", "+$12.4K / +$24.9K", "+$1.3K / +$2.5K"],
    ["3. Tighten travel, FICO < 640", "Risk model cut-off and limits cut segment volume −25% / −50%", "−$54K / −$107K", "−$3.8K / −$7.6K", "−$3.4K / −$6.9K"],
  ];
  body.forEach((r, i) => rows.push(r.map((c, j) => ({ text: c, options: {
    bold: j === 0, color: j >= 2 && c.startsWith("−") ? RED : INK, align: j >= 2 ? "center" : "left",
    fill: { color: i % 2 ? "F6F6FE" : WHITE } } }))));
  rows.push([
    { text: "Total vs 2025 baseline", options: { bold: true, color: INDIGO, fill: { color: LAV } } },
    { text: "Baseline: $2.85M GMV, $212K profit, 2.32% loss rate", options: { color: INK, fill: { color: LAV } } },
    { text: "+3.9% / +7.7%", options: { bold: true, color: INDIGO, align: "center", fill: { color: LAV } } },
    { text: "+7.4% / +14.8%", options: { bold: true, color: INDIGO, align: "center", fill: { color: LAV } } },
    { text: "Loss rate → 2.21% / 2.10%", options: { bold: true, color: GREEN, align: "center", fill: { color: LAV } } },
  ]);
  s.addTable(rows, { x: M, y: 1.55, w: W - 2 * M, colW: [2.9, 4.1, 1.75, 1.7, 1.683], rowH: [0.45, 0.8, 0.8, 0.8, 0.7],
    fontFace: FONT, fontSize: 12, border: { type: "solid", pt: 0.5, color: LINE }, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1] });

  card(s, M, 5.4, W - 2 * M, 1.2, "F6F6FE");
  pill(s, M + 0.25, 5.85, "NOT IN TOTALS", INDIGO, WHITE, 1.5);
  s.addText("4. Reprice 0% APR: raising the merchant fee +0.4 pt to interest-bearing parity adds ≈ +$1.9K a year (+0.9% profit), " +
    "if merchants keep the program — a call for Merchant Partnerships.",
    { x: M + 2.0, y: 5.5, w: W - 2 * M - 2.25, h: 0.95, fontFace: FONT, fontSize: 12, color: INK, margin: 0, valign: "middle", isTextBox: true });
  footnote(s, "One year = 2025 volume × seasoned margins/loss rates. The data is a sample of total volume, so read $ as relative sizes. New merchants assumed to perform like the vertical average.");
  s.addNotes("Each lever is sized on 2025 volume at seasoned margins, with a conservative and an ambitious case. Repeat customers are the biggest profit lever. Merchant acquisition adds steady growth: each vertical only has 3–4 merchants, so one new average merchant adds 25–35% to the vertical. Tightening travel costs some profit, which is a deliberate trade for lower risk. Net: profit +7–15%, GMV +4–8%, and the loss rate goes down. 0% APR repricing is on top, but depends on merchant response.");
}

// ---------- 7. Risks & next steps ----------
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  kicker(s, "Trade-offs and next steps", LAV2);
  title(s, "What could change the picture, and what we would do next", { color: WHITE });

  const cols = [
    ["Trade-offs & caveats", [
      "Tightening travel gives up ~$4–8K profit: the segment is still profitable on base assumptions; we trade it for lower loss and volatility",
      "0% APR repricing may push some merchants to drop the program",
      "Data is a synthetic sample; recent quarters aren't fully seasoned",
      "Delinquent-loan loss is an estimate (30% / 60% of unpaid principal)",
    ]],
    ["Watch", [
      "Split Pay loss rose from 1.1% (Q3-24) to 2.5–3.1% — its loans finish in 6 weeks, so this is a real trend",
      "Q4-2025 holiday vintage: last year's doubled its bad rate",
      "Travel data feed: all 160 corrupted duplicate loans were travel",
    ]],
    ["Next steps", [
      "Pilot the tighter travel cut-off on a share of traffic; track approvals and 90-day losses",
      "Test second-loan offers for big-ticket first-time customers",
      "Build a merchant pipeline for healthcare and home furnishing",
    ]],
  ];
  const cw = (W - 2 * M - 0.6) / 3;
  cols.forEach(([h, items], i) => {
    const x = M + i * (cw + 0.3);
    card(s, x, 1.6, cw, 3.7, "1C1E52");
    s.addText(h, { x: x + 0.25, y: 1.8, w: cw - 0.5, h: 0.45, fontFace: FONT, fontSize: 16, bold: true, color: WHITE, margin: 0, isTextBox: true });
    s.addText(items.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < items.length - 1 } })),
      { x: x + 0.25, y: 2.35, w: cw - 0.5, h: 2.85, fontFace: FONT, fontSize: 12.5, color: LAV2, paraSpaceAfter: 10, margin: 0, valign: "top", isTextBox: true });
  });
  s.addNotes("Be explicit about trade-offs: the travel tightening is a risk decision, not a profit one. Caveats on data. Three things to watch, notably Split Pay loss creeping up, which is the cleanest read on credit trend because those loans complete in six weeks. Next steps are pilots rather than big-bang changes.");
}

// ---------- Appendix A: data quality ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Appendix A · Data quality", MUTED);
  title(s, "One root cause explained most data issues; the rest were kept with stated assumptions");
  const H = (t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY } } });
  const rows = [[H("Issue"), H("Rows"), H("How handled")]];
  [
    ["Duplicate loan IDs; the extra copy has an impossible date (1923–25, 2031). All were travel loans", "160", "Dropped the bad copy — also fixed user-table mismatches and new-customer flags"],
    ["Missing FICO", "483 (6.8%)", "Not riskier than other loans; kept as an 'Unknown' band, no imputation"],
    ["credit_loss on delinquent loans is ~30% / ~60% of unpaid principal (100% for charge-offs)", "236", "Treated as expected loss (32% of loss $); stress case = all delinquents charge off"],
    ["Delinquency status earlier than loan age allows", "21", "Kept; can't tell whether date or flag is wrong; 0.3% of loans"],
    ["'current' Split Pay loans already fully repaid", "134", "Kept; no loss impact"],
    ["Split Pay term coded as 1 month (product is ~6 weeks)", "2,951", "Used 1.5 months where duration matters"],
  ].forEach((r, i) => rows.push(r.map((c, j) => ({ text: c, options: { color: INK, bold: j === 1, align: j === 1 ? "center" : "left",
    fill: { color: i % 2 ? "F6F6FE" : WHITE } } }))));
  s.addTable(rows, { x: M, y: 1.6, w: W - 2 * M, colW: [5.6, 1.4, 5.133], rowH: 0.62, fontFace: FONT, fontSize: 12,
    border: { type: "solid", pt: 0.5, color: LINE }, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1] });
  footnote(s, "Also checked and clean: keys and joins, product/term/APR rules, fee and repayment arithmetic. Code: notebooks/01_data_integrity.ipynb.");
}

// ---------- Appendix B: methodology ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Appendix B · Methodology", MUTED);
  title(s, "How the numbers were built");
  const blocks = [
    ["Unit economics per loan", "Contribution = merchant fee + interest collected − credit loss − funding cost. Funding = cost of funds × term × ½ principal (average balance of an amortizing loan). Full principal would lower margins ~1 pt but not change the vertical ranking."],
    ["Maturity", "Young loans haven't paid interest or defaulted yet. Loss/margin comparisons are checked on seasoned loans (originated by Jun 2025)."],
    ["Repeat customers", "First loans by Jun 2025; repeat = another loan within 180 days. Loss by new vs repeat checked on seasoned loans."],
    ["Sizing", "One year = 2025 volume × seasoned rates. Changes scale proportionally; new merchants perform like the vertical average."],
  ];
  blocks.forEach(([h, b], i) => {
    const x = M + (i % 2) * 6.2, y = 1.6 + Math.floor(i / 2) * 2.1;
    card(s, x, y, 5.9, 1.85, "F6F6FE");
    s.addText(h, { x: x + 0.3, y: y + 0.2, w: 5.3, h: 0.4, fontFace: FONT, fontSize: 15, bold: true, color: INDIGO, margin: 0, isTextBox: true });
    s.addText(b, { x: x + 0.3, y: y + 0.65, w: 5.3, h: 1.1, fontFace: FONT, fontSize: 12.5, color: INK, margin: 0, valign: "top", isTextBox: true });
  });
  footnote(s, "Code: notebooks/01_data_integrity.ipynb, 02_portfolio_risk.ipynb, 03_recommendation_sizing.ipynb.");
}

// ---------- Appendix C: margin by product over time ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Appendix C · Trend", MUTED);
  title(s, "Interest-bearing margin falls only because recent loans are young; Split Pay's dip is real");
  const labels = ["Q3-24", "Q4-24", "Q1-25", "Q2-25", "Q3-25", "Q4-25"];
  s.addChart(pres.charts.LINE, [
    { name: "Interest-bearing", labels, values: [13.2, 11.6, 11.6, 9.8, 9.1, 6.5] },
    { name: "0% APR", labels, values: [4.0, 3.0, 4.3, 4.0, 1.6, 3.7] },
    { name: "Split Pay", labels, values: [4.0, 1.7, 1.8, 2.4, 3.2, 2.8] },
  ], {
    x: M, y: 1.55, w: 7.6, h: 5.0, chartColors: [INDIGO, "1BAF7A", "EB6834"], lineSize: 2, lineDataSymbolSize: 7,
    showTitle: true, title: "Contribution margin by origination quarter (% of GMV)", titleFontFace: FONT, titleFontSize: 12, titleColor: INK,
    showLegend: true, legendPos: "b", legendFontFace: FONT, legendFontSize: 11,
    catAxisLabelColor: MUTED, valAxisLabelColor: MUTED, catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT,
    valAxisMinVal: 0, valGridLine: { color: "E6E6EE", size: 0.5 }, catGridLine: { style: "none" },
  });
  const px = 8.6, pw = W - M - px;
  s.addText([
    { text: "Interest-bearing: revenue falls 16.5% → 9.0% because recent loans haven't paid their interest yet; loss rate doesn't rise.", options: { bullet: true, breakLine: true } },
    { text: "Split Pay loans finish in ~6 weeks, so quarters are complete: loss rose from 1.1% to 2.5–3.1% and margin fell from 4.0% to ~2–3%.", options: { bullet: true, breakLine: true } },
    { text: "0% APR revenue is the merchant fee, fixed up front; margin ~3–4%.", options: { bullet: true } },
  ], { x: px, y: 1.7, w: pw, h: 4.6, fontFace: FONT, fontSize: 13, color: INK, paraSpaceAfter: 12, margin: 0, valign: "top", isTextBox: true });
}

pres.writeFile({ fileName: OUT }).then(f => console.log("wrote", f));
