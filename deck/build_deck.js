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

// ---------- 2. Portfolio health ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Portfolio health");
  title(s, "The book grew 33% year over year with stable credit quality and a ~7% margin");

  // Bars: GMV ($K, left axis). Lines: revenue, credit loss, funding cost (% of GMV, right axis)
  const qLabels = ["Q3-24", "Q4-24", "Q1-25", "Q2-25", "Q3-25", "Q4-25"];
  s.addChart([
    { type: pres.charts.BAR,
      data: [{ name: "GMV ($K, left axis)", labels: qLabels, values: [459, 664, 650, 704, 721, 771] }],
      options: { barDir: "col", chartColors: [LAV2], barGapWidthPct: 60,
        showValue: true, dataLabelPosition: "inEnd", dataLabelFontSize: 10, dataLabelColor: INK, dataLabelFontFace: FONT } },
    { type: pres.charts.LINE,
      data: [
        { name: "Revenue (% of GMV)", labels: qLabels, values: [10.8, 11.0, 11.3, 10.7, 9.7, 7.5] },
        { name: "Credit loss (% of GMV)", labels: qLabels, values: [1.4, 2.9, 2.1, 2.5, 2.0, 1.2] },
        { name: "Funding cost (% of GMV)", labels: qLabels, values: [1.0, 1.1, 1.1, 1.1, 1.1, 1.1] },
      ],
      options: { chartColors: [INDIGO, RED, MUTED], lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 7,
        secondaryValAxis: true, secondaryCatAxis: true } },
  ], {
    x: M, y: 1.45, w: 7.3, h: 4.9,
    showTitle: true, title: "GMV and unit economics by origination quarter", titleFontFace: FONT, titleFontSize: 13, titleColor: INK,
    showLegend: true, legendPos: "b", legendFontFace: FONT, legendFontSize: 10, legendColor: INK,
    valAxes: [
      { showValAxisTitle: true, valAxisTitle: "GMV ($K)", valAxisTitleFontSize: 10, valAxisTitleColor: MUTED,
        valAxisLabelColor: MUTED, valAxisLabelFontSize: 10, valAxisLabelFontFace: FONT, valAxisMinVal: 0, valAxisMaxVal: 900,
        valGridLine: { color: "E6E6EE", size: 0.5 } },
      { showValAxisTitle: true, valAxisTitle: "% of GMV", valAxisTitleFontSize: 10, valAxisTitleColor: MUTED,
        valAxisLabelColor: MUTED, valAxisLabelFontSize: 10, valAxisLabelFontFace: FONT, valAxisMinVal: 0, valAxisMaxVal: 12,
        valGridLine: { style: "none" } },
    ],
    catAxes: [
      { catAxisLabelColor: MUTED, catAxisLabelFontSize: 11, catAxisLabelFontFace: FONT },
      { catAxisHidden: true },
    ],
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
  footnote(s, "YoY = Jul–Dec 2025 vs Jul–Dec 2024. Q3–Q4 2025 loans are young, so their revenue and loss are still accruing. Funding = cost of funds × term × ½ principal.");
  s.addNotes("Healthy trajectory: GMV up every quarter, 33% year over year for the second half. Credit quality is stable: average FICO ~652 throughout. Economics: revenue 10% of GMV, loss 2%, funding 1%, leaving ~7% contribution. Two things to flag: growth is led by interest-bearing, which is good, but also by fashion, our lowest-margin vertical; and the Q4-2024 holiday vintage had double the bad rate. Recent quarters show lower margins only because those loans are young.");
}

// ---------- 3. Risk segmentation ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Where value and risk sit");
  title(s, "Travel carries the risk, fashion dilutes the margin, healthcare and home furnishing lead");

  const hdr = ["Vertical", "Share of loans", "Share of GMV", "Bad rate", "Loss rate", "Margin", "Share of profit"];
  const data = [
    ["Healthcare", "5%", "15%", "4%", "1.5%", "9.1%", "20%"],
    ["Travel", "6%", "20%", "21%", "3.5%", "8.7%", "24%"],
    ["Fitness", "8%", "7%", "5%", "1.2%", "7.4%", "7%"],
    ["Home furnishing", "8%", "14%", "4%", "1.2%", "7.4%", "15%"],
    ["Electronics", "16%", "20%", "4%", "1.4%", "6.4%", "18%"],
    ["Auto parts", "7%", "7%", "4%", "1.9%", "6.1%", "6%"],
    ["Fashion", "51%", "18%", "5%", "2.5%", "3.8%", "10%"],
  ];
  const highlight = { "Travel": { 3: RED, 4: RED }, "Fashion": { 1: INK, 4: RED, 5: RED }, "Healthcare": { 5: GREEN }, "Home furnishing": { 4: GREEN, 5: GREEN } };
  const rows = [hdr.map((h, j) => ({ text: h, options: { bold: true, color: WHITE, fill: { color: NAVY }, align: j ? "center" : "left" } }))];
  data.forEach((r, i) => rows.push(r.map((c, j) => {
    const col = (highlight[r[0]] || {})[j];
    return { text: c, options: { color: col || INK, bold: !!col || j === 0, align: j ? "center" : "left", fill: { color: i % 2 ? "F6F6FE" : WHITE } } };
  })));
  s.addTable(rows, { x: M, y: 1.6, w: 7.9, colW: [1.75, 1.05, 1.05, 0.95, 0.95, 0.95, 1.2], rowH: 0.52, fontFace: FONT, fontSize: 12,
    border: { type: "solid", pt: 0.5, color: LINE }, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1] });

  const px = 8.9, pw = W - M - px;
  const callouts = [
    ["Travel, FICO < 640", "~40% delinquent or charged off; 6.3% loss. Margin 6.7% falls to ~0% if delinquent loans charge off.", RED],
    ["Fashion", "51% of loans but 10% of profit. Small tickets; fashion Split Pay earns ~1% margin.", RED],
    ["Healthcare & home furnishing", "Low loss (1.2–1.5%) with high revenue. A healthcare customer is worth ~$215 in profit over their loans.", GREEN],
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

// ---------- 4. Product economics ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Where value sits by product");
  title(s, "Interest-bearing earns 82% of profit; 0% APR gives up interest without a higher merchant fee");

  const hdr = ["Product", "Share of loans", "Share of GMV", "Avg FICO", "Loss rate", "Merchant fee", "Interest", "Margin", "Share of profit"];
  const data = [
    ["Interest-bearing", "48%", "58%", "644", "2.3%", "6.1%", "7.4%", "9.9%", "82%"],
    ["0% APR", "11%", "18%", "683", "1.0%", "5.7%", "—", "3.4%", "9%"],
    ["Split Pay", "41%", "24%", "656", "2.2%", "5.0%", "—", "2.6%", "9%"],
  ];
  const highlight = { "Interest-bearing": { 6: GREEN, 7: GREEN, 8: GREEN }, "0% APR": { 3: GREEN, 4: GREEN, 5: RED, 7: RED },
                      "Split Pay": { 4: RED, 7: RED } };
  const rows = [hdr.map((h, j) => ({ text: h, options: { bold: true, color: WHITE, fill: { color: NAVY }, align: j ? "center" : "left" } }))];
  data.forEach((r, i) => rows.push(r.map((c, j) => {
    const col = (highlight[r[0]] || {})[j];
    return { text: c, options: { color: col || INK, bold: !!col || j === 0, align: j ? "center" : "left", fill: { color: i % 2 ? "F6F6FE" : WHITE } } };
  })));
  s.addTable(rows, { x: M, y: 1.6, w: W - 2 * M, colW: [2.0, 1.3, 1.3, 1.1, 1.2, 1.35, 1.2, 1.2, 1.483], rowH: 0.6, fontFace: FONT, fontSize: 13,
    border: { type: "solid", pt: 0.5, color: LINE }, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1] });

  const cards = [
    ["0% APR is under-priced", RED,
      "Lowest risk (FICO 683, 1.0% loss), but no interest and a 12–18 month term to fund. The promo is meant to be paid for by a higher merchant fee — yet at the same promo merchants its fee averages 5.0% vs 5.4% for interest-bearing."],
    ["Split Pay is an entry product", INDIGO,
      "Thin margin (2.6%) with loss as high as interest-bearing, but 63% of customers who start with Split Pay take an interest-bearing loan next."],
    ["Interest-bearing pays for the risk", GREEN,
      "Highest loss (2.3%) is more than covered by 7.4% interest income — 58% of GMV, 82% of profit."],
  ];
  const cw = (W - 2 * M - 0.6) / 3;
  cards.forEach(([h, col, b], i) => {
    const x = M + i * (cw + 0.3);
    card(s, x, 4.2, cw, 2.05, "F6F6FE");
    s.addText(h, { x: x + 0.25, y: 4.35, w: cw - 0.5, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: col, margin: 0, isTextBox: true });
    s.addText(b, { x: x + 0.25, y: 4.8, w: cw - 0.5, h: 1.35, fontFace: FONT, fontSize: 12, color: INK, margin: 0, valign: "top", isTextBox: true });
  });
  footnote(s, "Merchant fee and interest are % of GMV (dollar-weighted). Same-merchant comparison uses merchants with a promo program. Loss rate includes expected loss on delinquent loans.");
  s.addNotes("Same view by product. Interest-bearing is where the profit is: its loss rate is the highest, but interest income more than covers it. 0% APR is our safest product, yet it earns a third of the interest-bearing margin: we give up the interest, fund a 12-18 month loan, and the merchant fee is not higher — at the same merchants it is actually lower than on interest-bearing loans. That is the basis for repricing. Split Pay is thin too, but it is the entry product that leads customers to interest-bearing, so we keep it as an acquisition tool rather than reprice it.");
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
    ["86%", "of second loans stay in the same vertical — cross-selling is small (Appendix D)"],
  ];
  stats.forEach(([big, lab], i) => {
    const y = 1.55 + i * 1.22;
    card(s, px, y, pw, 1.08, LAV);
    s.addText(big, { x: px + 0.25, y: y + 0.08, w: pw - 0.5, h: 0.48, fontFace: FONT, fontSize: 24, bold: true, color: INDIGO, margin: 0, isTextBox: true });
    s.addText(lab, { x: px + 0.25, y: y + 0.56, w: pw - 0.5, h: 0.46, fontFace: FONT, fontSize: 11.5, color: INK, margin: 0, valign: "top", isTextBox: true });
  });
  footnote(s, "Customers whose first loan was by Jun 2025. Repeat rate is set by the vertical, not the product. Synthetic data: repeat timing is very regular, so read sizes as indicative.");
  s.addNotes("A customer who repays their first loan is a much better risk the second time: loss falls from 2.6% to 1.0%, and that holds in every vertical and product, including on seasoned loans. 82% of loss dollars come from first loans. Repeat rates are set by the vertical: frequent-purchase verticals like fashion and electronics already get ~95%; big-ticket verticals only ~20%, which is where the upside is. Split Pay works as an entry product: most customers move to interest-bearing on their next loan. And customers mostly come back to the same vertical (86%), so repeat growth has to be built vertical by vertical — cross-selling is small.");
}

// ---------- 6. Investment plan & sizing ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Investment plan");
  title(s, "Four moves add ~7–15% profit while lowering the loss rate");

  const H = (t, left) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, align: left ? "left" : "center" } });
  const rows = [[H("Move", true), H("Why", true), H("How (conservative / ambitious)", true), H("GMV"), H("Profit"), H("Credit loss $")]];
  const moves = [
    ["INVEST", GREEN, "Grow healthcare & home furnishing", "Best risk-adjusted margins (9.1%, 7.4%) with 1.2–1.5% loss",
      "Merchant acquisition: +10% / +20% vertical GMV (≈ 1 new merchant each at +20%)", "+$78K / +$155K", "+$7.0K / +$14.0K", "+$1.3K / +$2.7K"],
    ["INVEST", GREEN, "Repeat customers", "Repeat loans lose 2.5× less; 82% of loss comes from first loans",
      "+5 / +10 pts big-ticket repeat rate; 5% / 10% of repeat loans Split Pay → interest-bearing", "+$86K / +$172K", "+$12.4K / +$24.9K", "+$1.3K / +$2.5K"],
    ["TIGHTEN", RED, "Travel, FICO < 640", "~40% bad rate, 6.3% loss; margin ~0% if delinquents charge off",
      "Raise the risk model cut-off and lower limits: −25% / −50% segment volume", "−$54K / −$107K", "−$3.8K / −$7.6K", "−$3.4K / −$6.9K"],
    ["REPRICE", INDIGO, "0% APR merchant fee", "Fee is 5.0% vs 5.4% for interest-bearing at the same merchants",
      "+0.4 pt to parity (not in totals; depends on merchant response)", "—", "≈ +$1.9K", "—"],
  ];
  moves.forEach(([tag, col, name, why, how, gmv, profit, loss], i) => {
    const fill = { color: i % 2 ? "F6F6FE" : WHITE };
    // Two lines: conservative (top) / ambitious (bottom)
    const num = (t) => {
      const color = t.startsWith("−") ? RED : INK;
      const parts = t.split(" / ");
      const text = parts.length === 2
        ? [{ text: parts[0], options: { color, bold: true, breakLine: true } }, { text: parts[1], options: { color } }]
        : t;
      return { text, options: { align: "center", color, fill } };
    };
    rows.push([
      { text: [{ text: tag, options: { color: col, bold: true, fontSize: 9.5, breakLine: true } },
               { text: name, options: { color: INK, bold: true } }], options: { fill } },
      { text: why, options: { color: INK, fill } },
      { text: how, options: { color: INK, fill } },
      num(gmv), num(profit), num(loss),
    ]);
  });
  const T = (t, o = {}) => ({ text: t, options: { bold: true, color: INDIGO, align: "center", fill: { color: LAV }, ...o } });
  rows.push([
    T("Total (moves 1–3)", { align: "left" }),
    { text: "vs 2025 baseline: $2.85M GMV, $212K profit, 2.32% loss rate", options: { color: INK, fill: { color: LAV }, colspan: 2 } },
    T("+3.9% / +7.7%"), T("+7.4% / +14.8%"), T("Loss rate → 2.21% / 2.10%", { color: GREEN }),
  ]);
  s.addTable(rows, { x: M, y: 1.5, w: W - 2 * M, colW: [2.35, 2.9, 3.05, 1.25, 1.3, 1.283], rowH: [0.42, 0.88, 0.88, 0.88, 0.88, 0.62],
    fontFace: FONT, fontSize: 11, border: { type: "solid", pt: 0.5, color: LINE }, valign: "middle", margin: [0.04, 0.1, 0.04, 0.1] });

  s.addText("Profit grows faster than GMV and the loss rate falls: growth goes to low-loss verticals and repeat customers, and exposure is cut where risk is concentrated.",
    { x: M, y: 6.25, w: W - 2 * M, h: 0.5, fontFace: FONT, fontSize: 12, italic: true, color: MUTED, margin: 0, valign: "top", isTextBox: true });
  footnote(s, "GMV / profit / loss cells: top = conservative, bottom = ambitious. One year = 2025 volume × seasoned (Jul 2024–Jun 2025) rates; the data is a sample, so read $ as relative sizes.");
  s.addNotes("Four moves, each with the evidence and the sizing on one line. Two where we invest: healthcare and home furnishing merchants — one new average merchant adds 25–35% to the vertical — and repeat customers, which is the biggest profit lever. One where we pull back: travel below 640 FICO, through the risk model's cut-off and limits; it costs some profit, a deliberate trade for lower risk. One pricing fix: 0% APR merchant fee, shown separately because it depends on how merchants respond. Net: profit +7–15%, GMV +4–8%, and the loss rate goes down.");
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

// ---------- Appendix D: first vs second loan vertical ----------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, "Appendix D · Repeat customers", MUTED);
  title(s, "86% of repeat customers come back to the same vertical; fashion is the main source of cross-selling");
  const names = ["Fashion", "Electronics", "Fitness", "Home furnishing", "Travel", "Auto parts", "Healthcare"];
  const counts = [1340, 420, 202, 84, 49, 56, 37];
  const m = [[88, 3, 1, 2, 2, 2, 1], [9, 80, 2, 3, 1, 2, 1], [5, 5, 84, 1, 0, 3, 1], [6, 7, 2, 79, 2, 0, 4],
             [4, 2, 0, 2, 86, 4, 2], [11, 9, 0, 0, 0, 77, 4], [11, 3, 0, 0, 3, 8, 76]];
  const shade = (v) => v >= 50 ? [INDIGO, WHITE] : v >= 5 ? [LAV2, INK] : v >= 1 ? [LAV, INK] : [WHITE, MUTED];
  const hdr = [{ text: "First loan ↓ / second loan →", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
               { text: "Repeat customers", options: { bold: true, color: WHITE, fill: { color: NAVY }, align: "center" } },
               ...names.map(n => ({ text: n, options: { bold: true, color: WHITE, fill: { color: NAVY }, align: "center" } }))];
  const rows = [hdr];
  names.forEach((n, i) => rows.push([
    { text: n, options: { bold: true, color: INK } },
    { text: counts[i].toLocaleString("en-US"), options: { color: INK, align: "center" } },
    ...m[i].map(v => { const [f, c] = shade(v); return { text: v + "%", options: { align: "center", color: c, bold: v >= 50, fill: { color: f } } }; }),
  ]));
  s.addTable(rows, { x: M, y: 1.55, w: W - 2 * M, colW: [2.3, 1.3, 1.219, 1.219, 1.219, 1.219, 1.219, 1.219, 1.219], rowH: 0.55,
    fontFace: FONT, fontSize: 12, border: { type: "solid", pt: 0.5, color: LINE }, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1] });
  s.addText([
    { text: "Each row = customers whose first loan was in that vertical; cells = where their second loan was (row %).", options: { bullet: true, breakLine: true } },
    { text: "Fashion is the main cross-sell source: 157 fashion customers took their second loan elsewhere, 101 of them in big-ticket verticals.", options: { bullet: true } },
  ], { x: M, y: 6.1, w: W - 2 * M, h: 0.75, fontFace: FONT, fontSize: 12, color: INK, paraSpaceAfter: 4, margin: 0, valign: "top", isTextBox: true });
}

pres.writeFile({ fileName: OUT }).then(f => console.log("wrote", f));
