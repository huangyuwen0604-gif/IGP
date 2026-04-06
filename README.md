<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>IGP 個別輔導計畫批次生成系統</title>
<script src="https://unpkg.com/docx@8.5.0/build/index.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: Arial, '微軟正黑體', 'PingFang TC', sans-serif; background: #f0f2f5; color: #333; }

/* ─── 頂部 ─── */
.header { background: #2E75B6; color: white; padding: 14px 24px; }
.header h1 { font-size: 18px; font-weight: bold; }
.header p  { font-size: 12px; opacity: 0.8; margin-top: 3px; }

/* ─── 主體 ─── */
.container { max-width: 900px; margin: 0 auto; padding: 16px; }

/* ─── 分頁籤 ─── */
.tabs { display: flex; border-bottom: 2px solid #2E75B6; margin-bottom: 16px; }
.tab-btn { padding: 9px 20px; cursor: pointer; border: none; background: #e8edf3;
           font-size: 14px; font-family: inherit; border-radius: 6px 6px 0 0; margin-right: 4px;
           color: #555; transition: background 0.2s; }
.tab-btn.active { background: #2E75B6; color: white; font-weight: bold; }
.tab-pane { display: none; }
.tab-pane.active { display: block; }

/* ─── 卡片 ─── */
.card { background: white; border-radius: 8px; padding: 20px; margin-bottom: 14px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08); }
.card-title { font-size: 14px; font-weight: bold; color: #2E75B6;
              margin-bottom: 14px; padding-bottom: 8px; border-bottom: 1px solid #e0e8f0; }

/* ─── 表單欄位 ─── */
.form-row { display: flex; align-items: flex-start; margin-bottom: 12px; gap: 10px; }
.form-row label { min-width: 110px; font-size: 13px; padding-top: 7px; color: #444; }
.form-row input, .form-row select, .form-row textarea {
  flex: 1; padding: 7px 10px; border: 1px solid #ccd; border-radius: 5px;
  font-size: 13px; font-family: inherit; }
.form-row textarea { resize: vertical; line-height: 1.5; }
.radio-group { display: flex; gap: 16px; padding-top: 6px; }
.radio-group label { min-width: unset; padding-top: 0; cursor: pointer; }

/* ─── 週次表格 ─── */
.week-header { display: grid; grid-template-columns: 70px 1fr 1.5fr 36px;
               gap: 6px; padding: 7px 8px; background: #D9E1F2;
               border-radius: 5px 5px 0 0; font-size: 13px; font-weight: bold; color: #333; }
.week-row { display: grid; grid-template-columns: 70px 1fr 1.5fr 36px;
            gap: 6px; padding: 5px 8px; border-bottom: 1px solid #eee; align-items: center; }
.week-row input { padding: 5px 8px; border: 1px solid #ccd; border-radius: 4px;
                  font-size: 13px; font-family: inherit; width: 100%; }
.week-rows-container { border: 1px solid #D9E1F2; border-top: none; border-radius: 0 0 5px 5px; }
.btn-del { background: #C00000; color: white; border: none; border-radius: 4px;
           width: 28px; height: 28px; cursor: pointer; font-size: 14px; }
.btn-add { background: #4472C4; color: white; border: none; border-radius: 5px;
           padding: 7px 14px; cursor: pointer; font-size: 13px; margin-top: 8px; }

/* ─── 學生名單 ─── */
#students { width: 100%; height: 220px; padding: 10px; border: 1px solid #ccd;
            border-radius: 5px; font-size: 14px; line-height: 1.8; font-family: inherit; resize: vertical; }

/* ─── 底部操作 ─── */
.bottom-bar { position: sticky; bottom: 0; background: #f0f2f5; border-top: 1px solid #d0d8e0;
              padding: 12px 16px; display: flex; align-items: center; gap: 16px; flex-wrap: wrap; }
.bottom-bar label { font-size: 13px; color: #555; }
#outputNote { font-size: 12px; color: #888; }
.btn-generate { background: #70AD47; color: white; border: none; border-radius: 6px;
                padding: 11px 28px; font-size: 15px; font-weight: bold; cursor: pointer;
                font-family: inherit; margin-left: auto; transition: background 0.2s; }
.btn-generate:hover { background: #5a9036; }
.btn-generate:disabled { background: #aaa; cursor: not-allowed; }

/* ─── 提示訊息 ─── */
.notice { font-size: 12px; color: #888; margin-top: 6px; }
.hint { color: #2E75B6; font-size: 12px; padding: 6px 10px; background: #e8f0fb;
        border-radius: 4px; margin-top: 8px; }
</style>
</head>
<body>

<div class="header">
  <h1>IGP 個別輔導計畫批次生成系統</h1>
  <p>新竹縣資賦優異學生個別輔導計畫 ── 一次輸入課程資訊，自動產生所有學生文件</p>
</div>

<div class="container">

  <!-- ── 分頁籤 ── -->
  <div class="tabs">
    <button class="tab-btn active" onclick="switchTab('course')">📋 課程資訊</button>
    <button class="tab-btn" onclick="switchTab('weeks')">📅 週次安排</button>
    <button class="tab-btn" onclick="switchTab('students')">👩‍🎓 學生名單</button>
  </div>

  <!-- ═══ Tab 1：課程資訊 ═══ -->
  <div class="tab-pane active" id="tab-course">
    <div class="card">
      <div class="card-title">基本資訊（所有學生共用）</div>

      <div class="form-row">
        <label>學年度學期</label>
        <input type="text" id="semester" value="113學年度第1學期" style="max-width:200px">
      </div>
      <div class="form-row">
        <label>教育階段</label>
        <div class="radio-group">
          <label><input type="radio" name="stage" value="高中"> 高中</label>
          <label><input type="radio" name="stage" value="國中" checked> 國中</label>
          <label><input type="radio" name="stage" value="國小"> 國小</label>
          <label><input type="radio" name="stage" value="學前"> 學前</label>
        </div>
      </div>
      <div class="form-row">
        <label>課程類型</label>
        <div class="radio-group">
          <label><input type="radio" name="ctype" value="領域學習課程" checked> 領域學習課程</label>
          <label><input type="radio" name="ctype" value="特殊需求課程"> 特殊需求課程</label>
        </div>
      </div>
      <div class="form-row">
        <label>課程名稱</label>
        <input type="text" id="courseName" placeholder="例：科學資優課程">
      </div>
      <div class="form-row">
        <label>教學年級/組別</label>
        <input type="text" id="grade" placeholder="例：七年級A組" style="max-width:200px">
      </div>
      <div class="form-row">
        <label>教學節數/週</label>
        <input type="text" id="hours" placeholder="例：2節" style="max-width:100px">
      </div>
      <div class="form-row">
        <label>教學者</label>
        <input type="text" id="teacher" placeholder="老師姓名" style="max-width:160px">
      </div>
    </div>

    <div class="card">
      <div class="card-title">課程目標（從課程計畫複製貼上即可）</div>
      <div class="form-row">
        <label>學年目標</label>
        <textarea id="annualGoals" rows="5" placeholder="整個學期的課程目標，5條以內&#10;1. &#10;2. &#10;3. "></textarea>
      </div>
      <div class="form-row">
        <label>核心素養</label>
        <textarea id="coreComp" rows="3" placeholder="參考課程計畫填入，或留空"></textarea>
      </div>
      <div class="form-row">
        <label>學習表現</label>
        <textarea id="learningPerf" rows="3" placeholder="參考課程計畫填入，或留空"></textarea>
      </div>
      <div class="form-row">
        <label>學習內容</label>
        <textarea id="learningContent" rows="3" placeholder="參考課程計畫填入，或留空"></textarea>
      </div>
    </div>
  </div>

  <!-- ═══ Tab 2：週次安排 ═══ -->
  <div class="tab-pane" id="tab-weeks">
    <div class="card">
      <div class="card-title">週次與課程內容（所有學生共用）</div>
      <div class="week-header">
        <span>周次</span>
        <span>課程/單元名稱</span>
        <span>教學重點（從課程計畫複製）</span>
        <span></span>
      </div>
      <div class="week-rows-container" id="weekRows"></div>
      <button class="btn-add" onclick="addWeekRow()">＋ 新增週次</button>
      <p class="notice">提示：週次可填 1-4、5-8 等範圍，教學重點留空也沒關係，文件中會保留空格</p>
    </div>
  </div>

  <!-- ═══ Tab 3：學生名單 ═══ -->
  <div class="tab-pane" id="tab-students">
    <div class="card">
      <div class="card-title">學生姓名清單</div>
      <textarea id="students" placeholder="每行輸入一位學生姓名&#10;&#10;王小明&#10;李大華&#10;張志明&#10;&#10;（或用逗號分隔：王小明,李大華,張志明）"></textarea>
      <div class="hint">💡 每位學生將產生一份獨立的 IGP Word 文件（.docx），最後打包成 ZIP 壓縮檔下載</div>
    </div>
  </div>

</div><!-- end container -->

<!-- ── 底部操作列 ── -->
<div class="bottom-bar">
  <span id="outputNote">填寫完成後，點右側按鈕批次下載所有 IGP 文件</span>
  <button class="btn-generate" id="genBtn" onclick="generateAll()">
    🚀 批次生成並下載 ZIP
  </button>
</div>

<script>
// ══════════════════════════════════════════════
//  分頁籤切換
// ══════════════════════════════════════════════
function switchTab(name) {
  document.querySelectorAll('.tab-btn').forEach((b, i) => {
    b.classList.toggle('active', ['course','weeks','students'][i] === name);
  });
  document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
}

// ══════════════════════════════════════════════
//  週次列管理
// ══════════════════════════════════════════════
let weekCount = 0;

function addWeekRow(week = '', unit = '', point = '') {
  const id = ++weekCount;
  const div = document.createElement('div');
  div.className = 'week-row';
  div.id = 'week-' + id;
  div.innerHTML = `
    <input type="text" value="${esc(week)}" placeholder="1-4">
    <input type="text" value="${esc(unit)}" placeholder="單元名稱">
    <input type="text" value="${esc(point)}" placeholder="教學重點">
    <button class="btn-del" onclick="removeWeek(${id})" title="刪除">✕</button>
  `;
  document.getElementById('weekRows').appendChild(div);
}

function removeWeek(id) {
  const el = document.getElementById('week-' + id);
  if (el) el.remove();
}

function esc(s) {
  return (s || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;');
}

function getWeeks() {
  const rows = document.querySelectorAll('.week-row');
  const result = [];
  rows.forEach(row => {
    const inputs = row.querySelectorAll('input');
    result.push({
      week:  inputs[0].value.trim(),
      unit:  inputs[1].value.trim(),
      point: inputs[2].value.trim(),
    });
  });
  return result;
}

// Default weeks
['1-4','5-8','9-12','13-16','17-20'].forEach(w => addWeekRow(w));

// ══════════════════════════════════════════════
//  表單資料收集
// ══════════════════════════════════════════════
function getFormData() {
  return {
    semester:       document.getElementById('semester').value.trim(),
    stage:          document.querySelector('input[name="stage"]:checked')?.value || '國中',
    courseType:     document.querySelector('input[name="ctype"]:checked')?.value || '領域學習課程',
    courseName:     document.getElementById('courseName').value.trim(),
    grade:          document.getElementById('grade').value.trim(),
    hours:          document.getElementById('hours').value.trim(),
    teacher:        document.getElementById('teacher').value.trim(),
    annualGoals:    document.getElementById('annualGoals').value.trim(),
    coreComp:       document.getElementById('coreComp').value.trim(),
    learningPerf:   document.getElementById('learningPerf').value.trim(),
    learningContent:document.getElementById('learningContent').value.trim(),
    weeks:          getWeeks(),
  };
}

function parseStudents() {
  const raw = document.getElementById('students').value;
  return raw
    .replace(/，/g, ',').replace(/、/g, ',')
    .split(/[\n,]/)
    .map(s => s.trim())
    .filter(s => s.length > 0);
}

// ══════════════════════════════════════════════
//  DOCX 生成
// ══════════════════════════════════════════════
function buildDoc(studentName, d) {
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  } = window.docx;

  const FONT = '標楷體';
  const W    = 9639;
  const COLS = [700, 1800, 2000, 2039, 1500, 1600];

  const bdr  = { style: BorderStyle.SINGLE, size: 1, color: '000000' };
  const bdrs = { top: bdr, bottom: bdr, left: bdr, right: bdr };
  const marg = { top: 60, bottom: 60, left: 100, right: 100 };

  const sumCols = (from, n) => COLS.slice(from, from + n).reduce((a,b) => a+b, 0);

  // ─ TextRun helper ─
  function tx(text, opts = {}) {
    return new TextRun({
      text: text || '',
      font: FONT,
      size: (opts.size || 10) * 2,   // half-points
      bold: opts.bold || false,
      underline: opts.underline ? {} : undefined,
    });
  }

  // ─ Paragraphs from multi-line string ─
  function paras(text, opts = {}) {
    return (text || '').split('\n').map(line =>
      new Paragraph({
        alignment: opts.align || AlignmentType.LEFT,
        spacing: { before: 0, after: 0, line: 276 },
        children: [tx(line, opts)],
      })
    );
  }

  // ─ TableCell helper ─
  function tc(text, { width, cs = 1, rs = 1, bg, bold = false,
                       align = AlignmentType.LEFT, size = 10 } = {}) {
    const children = Array.isArray(text) ? text : paras(text, { bold, align, size });
    return new TableCell({
      borders: bdrs,
      width: { size: width, type: WidthType.DXA },
      columnSpan: cs,
      rowSpan: rs,
      margins: marg,
      verticalAlign: VerticalAlign.CENTER,
      shading: bg ? { fill: bg, type: ShadingType.CLEAR, color: 'auto' } : undefined,
      children,
    });
  }

  const H  = 'BDD7EE'; // header blue
  const L  = 'D9E1F2'; // label blue
  const SB = 'EBF3FB'; // sub blue

  // Checkbox helpers
  const isLY = d.courseType.includes('領域');
  const stageBoxes = ['高中','國中','國小','學前']
    .map(s => (s === d.stage ? '■' : '□') + s).join('　');

  // ─ Weekly rows ─
  const ADJ = '□加深□加廣□重組\n□統整教學主題\n□其他：＿＿＿';
  const weekRows = (d.weeks || []).map(w => new TableRow({ children: [
    tc(w.week,  { width: COLS[0], align: AlignmentType.CENTER }),
    tc(w.unit,  { width: COLS[1] }),
    tc(w.point, { width: COLS[2] }),
    tc(ADJ,     { width: COLS[3] }),
    tc('',      { width: COLS[4] }),
    tc('',      { width: COLS[5] }),
  ]}));

  const mainTable = new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: COLS,
    rows: [
      // Row 0: semester title
      new TableRow({ children: [
        tc((d.semester||'') + '　目標',
           { width: W, cs: 6, bg: H, bold: true, size: 11, align: AlignmentType.CENTER }),
      ]}),
      // Row 1: column headers
      new TableRow({ children: [
        tc('課程類型',       { width: COLS[0], bg: L, bold: true, size: 10, align: AlignmentType.CENTER }),
        tc('課程名稱',       { width: COLS[1], bg: L, bold: true, size: 10, align: AlignmentType.CENTER }),
        tc('教學年級\n/組別',{ width: COLS[2], bg: L, bold: true, size: 10, align: AlignmentType.CENTER }),
        tc('教學節數/週',   { width: COLS[3], bg: L, bold: true, size: 10, align: AlignmentType.CENTER }),
        tc('教學者',         { width: sumCols(4,2), cs: 2, bg: L, bold: true, size: 10, align: AlignmentType.CENTER }),
      ]}),
      // Row 2: course data
      new TableRow({ children: [
        tc((isLY?'■':'□')+'領域學習課程\n'+(isLY?'□':'■')+'特殊需求課程',
           { width: COLS[0] }),
        tc(d.courseName || '', { width: COLS[1], align: AlignmentType.CENTER }),
        tc(d.grade      || '', { width: COLS[2], align: AlignmentType.CENTER }),
        tc(d.hours      || '', { width: COLS[3], align: AlignmentType.CENTER }),
        tc(d.teacher    || '', { width: sumCols(4,2), cs: 2, align: AlignmentType.CENTER }),
      ]}),
      // Rows 3-6: info rows
      ...[ ['學年目標', d.annualGoals],
           ['核心素養', d.coreComp],
           ['學習表現', d.learningPerf],
           ['學習內容', d.learningContent],
      ].map(([lbl, val]) => new TableRow({ children: [
        tc(lbl, { width: COLS[0], bg: L, bold: true, size: 10, align: AlignmentType.CENTER }),
        tc(val || '', { width: sumCols(1,5), cs: 5 }),
      ]})),
      // Row 7: adjustment – content
      new TableRow({ children: [
        tc('課程調整\n策略', { width: COLS[0], bg: L, bold: true, size: 10,
                               align: AlignmentType.CENTER, rs: 2 }),
        tc('學習內容', { width: COLS[1], bg: SB, bold: true, size: 10, align: AlignmentType.CENTER }),
        tc('□重組　□加深　□加廣　□濃縮　□加速\n□跨領域/科目統整教學主題\n□其他：＿＿＿＿＿＿',
           { width: sumCols(2,4), cs: 4 }),
      ]}),
      // Row 8: adjustment – process  (col 0 is covered by rowSpan above)
      new TableRow({ children: [
        tc('學習歷程', { width: COLS[1], bg: SB, bold: true, size: 10, align: AlignmentType.CENTER }),
        tc('□高層次思考　□開放式問題　□發現式學習　□推理的證據\n□選擇的自由　□團體式的互動　□彈性的教學進度\n□多樣性的歷程\n□其他：＿＿＿＿',
           { width: sumCols(2,4), cs: 4 }),
      ]}),
      // Row 9: week headers
      new TableRow({ children: [
        tc('周次',       { width: COLS[0], bg: L, bold: true, align: AlignmentType.CENTER }),
        tc('課程內容',   { width: COLS[1], bg: L, bold: true, align: AlignmentType.CENTER }),
        tc('教學重點',   { width: COLS[2], bg: L, bold: true, align: AlignmentType.CENTER }),
        tc('學習內容調整',{ width: COLS[3], bg: L, bold: true, align: AlignmentType.CENTER }),
        tc('評量\n方式', { width: COLS[4], bg: L, bold: true, align: AlignmentType.CENTER }),
        tc('評量日期\n評量結果', { width: COLS[5], bg: L, bold: true, align: AlignmentType.CENTER }),
      ]}),
      ...weekRows,
    ],
  });

  const legendTable = new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [1500, W - 1500],
    rows: [
      new TableRow({ children: [
        tc('評量方式', { width: 1500, bg: L, bold: true, align: AlignmentType.CENTER }),
        tc('1.口頭發表　2.書面報告　3.作業單　4.器材操作　5.成品製作　6.活動設計\n7.觀察評量　8.演示評量　9.檔案評量　10.其他：＿＿＿＿＿',
           { width: W - 1500 }),
      ]}),
      new TableRow({ children: [
        tc('評量結果', { width: 1500, bg: L, bold: true, align: AlignmentType.CENTER }),
        tc('特優 100～95%　優 94～85%　良 84～70%　中等 69～55%　中下 54～40%　待加強 39%以下',
           { width: W - 1500 }),
      ]}),
    ],
  });

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 },
        },
      },
      children: [
        // Title
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 80 },
          children: [tx('新竹縣資賦優異學生個別輔導計畫（IGP）', { bold: true, size: 14 })],
        }),
        // Reference
        new Paragraph({
          spacing: { before: 0, after: 80 },
          children: [tx('參考資優學生個別輔導計畫（IGP）線上版參考格式（國立臺灣師範大學郭靜姿教授計畫提供）', { size: 9 })],
        }),
        // Student line
        new Paragraph({
          spacing: { before: 0, after: 80 },
          children: [
            tx('學生：', { size: 11 }),
            tx(studentName, { size: 11, bold: true, underline: true }),
            tx('　　教育階段：' + stageBoxes, { size: 11 }),
          ],
        }),
        // Section title
        new Paragraph({
          spacing: { before: 100, after: 80 },
          children: [tx('五、教育目標及課程調整', { bold: true, size: 12 })],
        }),
        mainTable,
        new Paragraph({
          spacing: { before: 140, after: 60 },
          children: [tx('備註：', { bold: true })],
        }),
        legendTable,
      ],
    }],
  });

  return doc;
}

// ══════════════════════════════════════════════
//  批次生成 + 下載 ZIP
// ══════════════════════════════════════════════
async function generateAll() {
  const students = parseStudents();
  if (!students.length) {
    alert('請先在「學生名單」頁填入學生姓名！');
    switchTab('students');
    return;
  }

  const btn = document.getElementById('genBtn');
  const note = document.getElementById('outputNote');
  btn.disabled = true;
  btn.textContent = '⏳ 生成中…';

  const d = getFormData();
  const zip = new JSZip();

  try {
    for (let i = 0; i < students.length; i++) {
      const name = students[i];
      note.textContent = `⏳ 正在生成 ${name}（${i+1}/${students.length}）…`;
      const doc = buildDoc(name, d);
      const buf = await window.docx.Packer.toBuffer(doc);
      zip.file(`IGP_${name}.docx`, buf);
      await new Promise(r => setTimeout(r, 30)); // let UI breathe
    }

    note.textContent = `✅ 打包中…`;
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    saveAs(zipBlob, `IGP_批次文件_${d.semester || ''}.zip`);
    note.textContent = `✅ 已下載 ${students.length} 份 IGP 文件`;
  } catch (err) {
    console.error(err);
    alert('生成發生錯誤：\n' + err.message);
    note.textContent = '❌ 生成失敗，請查看主控台';
  } finally {
    btn.disabled = false;
    btn.textContent = '🚀 批次生成並下載 ZIP';
  }
}
</script>
</body>
</html>
