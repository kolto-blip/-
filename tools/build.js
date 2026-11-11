// Іске қосу: node tools/build.js  → word/ және google-forms/ файлдарын жасайды
const fs = require("fs"), path = require("path");
const D = require("./data");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, ShadingType, PageBreak, HeadingLevel } = require("docx");
const root = path.join(__dirname, "..");
const L = ["A", "B", "C", "D"];
const F = "Times New Roman";
const p = (text, o = {}) => new Paragraph({ spacing: { after: o.after ?? 80 }, alignment: o.align,
  keepNext: o.keepNext, indent: o.indent, children: [new TextRun({ text, font: F, size: o.size ?? 24, bold: o.bold })] });
const border = { style: BorderStyle.SINGLE, size: 4, color: "808080" };
const borders = { top: border, bottom: border, left: border, right: border };
function table(rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map((r, ri) => new TableRow({ children: r.map((c, ci) => new TableCell({
      borders, width: { size: widths[ci], type: WidthType.DXA },
      shading: ri === 0 ? { type: ShadingType.CLEAR, fill: "E7EEF7", color: "auto" } : undefined,
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
      children: [p(c, { after: 0, align: AlignmentType.CENTER, bold: ri === 0, size: 22 })] })) })) });
}
const body = [
  p(D.title, { size: 32, bold: true, align: AlignmentType.CENTER, after: 60 }),
  p(D.variant, { size: 28, bold: true, align: AlignmentType.CENTER, after: 200 }),
  p("Аты-жөні: ______________________________   Сынып: ________   Күні: ____________", { after: 200 }),
  p("Тапсырма саны — 10. Әр тапсырма — 1 балл (барлығы 10 балл). Әр тапсырмада бір ғана дұрыс жауап бар. Ұсынылатын уақыт — 20 минут.", { after: 240, size: 22 }),
];
D.questions.forEach((x, i) => {
  body.push(new Paragraph({ keepNext: true, spacing: { before: 200, after: 80 }, children: [
    new TextRun({ text: `${i + 1}. ${x.h}. `, font: F, size: 24, bold: true }),
    new TextRun({ text: x.q, font: F, size: 24 })] }));
  if (x.table) { body.push(table(x.table, [2200, 1350, 1350, 1350, 1350, 1350])); body.push(p("", { after: 60 })); }
  x.o.forEach((o, j) => body.push(p(`${L[j]}) ${o}`, { indent: { left: 400 }, after: 20, keepNext: j < 3 })));
});
body.push(new Paragraph({ children: [new PageBreak()] }));
body.push(p("Жауаптар кілті", { size: 28, bold: true, align: AlignmentType.CENTER, after: 160 }));
body.push(table([["№", ...D.questions.map((_, i) => String(i + 1))], ["Жауап", ...D.questions.map(x => L[x.a])]],
  [1300, ...Array(10).fill(780)]));
body.push(p("", { after: 120 }));
body.push(p("Қысқаша шешімдер", { size: 26, bold: true, after: 120 }));
D.questions.forEach((x, i) => body.push(p(`${i + 1}. ${x.s}  Жауабы: ${L[x.a]}.`, { size: 22, after: 60 })));
const doc = new Document({ sections: [{ properties: { page: { margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } }, children: body }] });
fs.mkdirSync(path.join(root, "word"), { recursive: true });
Packer.toBuffer(doc).then(b => fs.writeFileSync(path.join(root, "word", "matematikalyq-sauattylyq-nusqa-1.docx"), b));

// Google Apps Script — Google Forms тестін жасайды
const qs = D.questions.map((x, i) => {
  let help = x.q;
  if (x.table) help += "\n\n" + x.table.map(r => r.join(" | ")).join("\n");
  return `  {title: ${JSON.stringify(`${i + 1}. ${x.h}`)}, help: ${JSON.stringify(help)}, options: ${JSON.stringify(x.o)}, correct: ${x.a}, solution: ${JSON.stringify(x.s)}}`;
}).join(",\n");
const gs = `/**
 * ${D.title} — ${D.variant}: Google Forms тестін автоматты түрде жасайды.
 *
 * Қалай қолдану керек:
 * 1. https://script.google.com сайтына кіріп, «Жаңа жоба» (New project) басыңыз.
 * 2. Бар кодты өшіріп, осы файлдың бүкіл мазмұнын қойыңыз.
 * 3. Жоғарыдан createUbtForm функциясын таңдап, «Іске қосу» (Run) басыңыз.
 * 4. Google рұқсат сұрайды — рұқсат беріңіз.
 * 5. «Орындау журналында» (Execution log) форманың сілтемелері шығады.
 *    Форма сіздің Google Drive-та да пайда болады.
 */
const QUESTIONS = [
${qs}
];

function createUbtForm() {
  const form = FormApp.create(${JSON.stringify(`${D.title} — ${D.variant}`)});
  form.setIsQuiz(true)
      .setDescription("Тапсырма саны — 10. Әр тапсырма — 1 балл (барлығы 10 балл). Әр тапсырмада бір ғана дұрыс жауап бар. Ұсынылатын уақыт — 20 минут.")
      .setCollectEmail(false)
      .setProgressBar(true)
      .setShuffleQuestions(false);
  form.addTextItem().setTitle("Аты-жөні").setRequired(true);
  form.addTextItem().setTitle("Сынып").setRequired(true);
  const letters = ["A", "B", "C", "D"];
  QUESTIONS.forEach(function (q) {
    const item = form.addMultipleChoiceItem();
    item.setTitle(q.title)
        .setHelpText(q.help)
        .setChoices(q.options.map(function (o, j) { return item.createChoice(letters[j] + ") " + o, j === q.correct); }))
        .setPoints(1)
        .setRequired(true);
    item.setFeedbackForIncorrect(FormApp.createFeedback().setText("Шешімі: " + q.solution).build());
    item.setFeedbackForCorrect(FormApp.createFeedback().setText("Дұрыс! " + q.solution).build());
  });
  Logger.log("Өңдеу сілтемесі: " + form.getEditUrl());
  Logger.log("Оқушыларға сілтеме: " + form.getPublishedUrl());
}
`;
fs.mkdirSync(path.join(root, "google-forms"), { recursive: true });
fs.writeFileSync(path.join(root, "google-forms", "create-form-nusqa-1.gs"), gs);
