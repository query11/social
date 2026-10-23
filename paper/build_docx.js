// 把 draft.md 排版成符合课程“技术规范”的 Word 文件：
// A4；标题小二黑体；姓名学号小四仿宋；正文小四宋体；全文 1.5 倍行距；脚注注释，无文末参考文献。
// 用法：node paper/build_docx.js <稿件.md> [输出路径]，默认输出为同目录下“<标题>（初稿）.docx”
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, FootnoteReferenceRun, Footer,
  PageNumber, AlignmentType, LineRuleType,
} = require("docx");

const SRC = process.argv[2];
if (!SRC) throw new Error("用法：node paper/build_docx.js <稿件.md> [输出路径]");

// 字号（半磅）：小二 18pt、四号 14pt、小四 12pt、小五 9pt
const SIZE = { xiaoEr: 36, siHao: 28, xiaoSi: 24, xiaoWu: 18 };
const LINE_1_5 = { line: 360, lineRule: LineRuleType.AUTO };
const INDENT_2_CHARS = 480; // 小四两个汉字宽

const hasCJK = (s) => /[　-〿一-鿿＀-￯]/.test(s);
// hint=eastAsia 让引号、破折号、省略号、间隔号按中文全角字形显示
const font = (eastAsia, text = "中") => ({
  ascii: "Times New Roman", hAnsi: "Times New Roman", cs: "Times New Roman",
  eastAsia, ...(hasCJK(text) ? { hint: "eastAsia" } : {}),
});
const run = (text, eastAsia, size, extra = {}) =>
  new TextRun({ text, font: font(eastAsia, text), size, ...extra });

// 把 *斜体* 片段拆成多个 run（脚注里的书名、刊名）
function inlineRuns(text, eastAsia, size) {
  return text.split(/(\*[^*]+\*)/).filter(Boolean).map((part) =>
    part.startsWith("*")
      ? run(part.slice(1, -1), eastAsia, size, { italics: true })
      : run(part, eastAsia, size));
}

const lines = fs.readFileSync(SRC, "utf8").split(/\r?\n/);
const title = lines.find((l) => l.startsWith("# ")).slice(2).trim();
const OUT = process.argv[3] || path.join(path.dirname(SRC), `${title}（初稿）.docx`);
const footnotes = {};
for (const l of lines) {
  const m = l.match(/^\[\^(\d+)\]:\s*(.+)$/);
  if (m) {
    footnotes[m[1]] = {
      children: [new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        children: inlineRuns(m[2], "宋体", SIZE.xiaoWu),
      })],
    };
  }
}

const children = [];
for (const l of lines) {
  if (!l.trim() || /^\[\^\d+\]:/.test(l)) continue;
  let m;
  if ((m = l.match(/^# (.+)$/))) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { ...LINE_1_5, after: 60 },
      children: [run(m[1], "黑体", SIZE.xiaoEr)],
    }));
  } else if ((m = l.match(/^## (.+)$/))) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { ...LINE_1_5, after: 120 },
      children: [run(m[1], "楷体", SIZE.siHao)],
    }));
  } else if (l.startsWith("姓名")) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { ...LINE_1_5, after: 240 },
      children: [run(l, "仿宋", SIZE.xiaoSi)],
    }));
  } else if ((m = l.match(/^### (.+)$/))) {
    children.push(new Paragraph({
      outlineLevel: 0, keepNext: true, spacing: { ...LINE_1_5, before: 120, after: 60 },
      children: [run(m[1], "黑体", SIZE.xiaoSi)],
    }));
  } else if ((m = l.match(/^#### (.+)$/))) {
    children.push(new Paragraph({
      outlineLevel: 1, keepNext: true, spacing: LINE_1_5,
      indent: { firstLine: INDENT_2_CHARS },
      children: [run(m[1], "宋体", SIZE.xiaoSi, { bold: true })],
    }));
  } else {
    const parts = l.split(/(\[\^\d+\])/).filter(Boolean);
    children.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED, spacing: LINE_1_5,
      indent: { firstLine: INDENT_2_CHARS },
      children: parts.map((p) => {
        const ref = p.match(/^\[\^(\d+)\]$/);
        return ref ? new FootnoteReferenceRun(Number(ref[1])) : run(p, "宋体", SIZE.xiaoSi);
      }),
    }));
  }
}

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: font("宋体"), size: SIZE.xiaoSi },
        paragraph: { spacing: LINE_1_5 },
      },
      footnoteText: {
        run: { font: font("宋体"), size: SIZE.xiaoWu },
        paragraph: { spacing: { line: 240, lineRule: LineRuleType.AUTO } },
      },
    },
  },
  footnotes,
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1440, bottom: 1440, left: 1800, right: 1800 },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], font: font("宋体"), size: SIZE.xiaoWu })],
        })],
      }),
    },
    children,
  }],
});

// docx-js 不支持脚注编号格式，打包后补上：①②③ 编号、全文连续编号
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("docx")] }));
const FOOTNOTE_PR =
  '<w:footnotePr><w:numFmt w:val="decimalEnclosedCircleChinese"/><w:numRestart w:val="continuous"/></w:footnotePr>';

Packer.toBuffer(doc)
  .then(JSZip.loadAsync)
  .then(async (zip) => {
    const xml = await zip.file("word/document.xml").async("string");
    const patched = xml.replace(/(<w:sectPr>(?:<w:(?:header|footer)Reference[^>]*\/>)*)/, `$1${FOOTNOTE_PR}`);
    if (patched === xml) throw new Error("未找到 sectPr，脚注格式未写入");
    zip.file("word/document.xml", patched);
    fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
    console.log("wrote", OUT);
  });
