Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("translateBtn").onclick = handleTranslate;
  }
});

// Konversi HANYA teks alfabet ke karakter miring Unicode
function toUnicodeItalic(text) {
  return text.split('').map(char => {
    const code = char.charCodeAt(0);
    // Huruf besar A-Z (A=0x1D434)
    if (code >= 65 && code <= 90) {
      return String.fromCodePoint(0x1D434 + (code - 65));
    }
    // Huruf kecil a-z (a=0x1D44E)
    if (code >= 97 && code <= 122) {
      if (char === 'h') return 'ℎ'; // Pengecualian standar Unicode
      return String.fromCodePoint(0x1D44E + (code - 97));
    }
    // Angka, spasi, dan simbol tetap standar
    return char;
  }).join('');
}

// Teks yang sudah berisi huruf miring Unicode dianggap sudah diterjemahkan
const ITALIC_MARK = /[\u{1D434}-\u{1D467}]/u;
const HAS_LETTER = /[A-Za-zÀ-ɏ]/;

function isTranslatable(value) {
  return typeof value === "string"
    && value.trim() !== ""
    && HAS_LETTER.test(value)
    && !ITALIC_MARK.test(value);
}

// MyMemory membatasi 500 karakter per request, jadi teks panjang dipecah per kalimat
function splitChunks(text, max = 450) {
  if (text.length <= max) return [text];
  const sentences = text.match(/[^.!?\n]+[.!?\n]*\s*/g) || [text];
  const chunks = [];
  let current = "";
  for (const s of sentences) {
    if ((current + s).length > max && current) {
      chunks.push(current);
      current = "";
    }
    current += s;
  }
  if (current) chunks.push(current);
  return chunks.flatMap(c => c.length > max ? c.match(new RegExp(`.{1,${max}}`, "gs")) : [c]);
}

const translationCache = new Map();

// Fetch terjemahan ID -> EN
async function fetchTranslation(text) {
  const url = `https://api.mymemory.translated.net/get?q=${encodeURIComponent(text)}&langpair=id|en`;
  const response = await fetch(url);
  const data = await response.json();

  if (data && data.responseData && data.responseData.translatedText) {
    return data.responseData.translatedText;
  }
  throw new Error("Gagal mengambil data terjemahan.");
}

async function translateCached(text) {
  const key = text.trim();
  if (translationCache.has(key)) return translationCache.get(key);
  const parts = [];
  for (const chunk of splitChunks(key)) {
    parts.push((await fetchTranslation(chunk.trim())).trim());
  }
  const result = parts.join(" ");
  translationCache.set(key, result);
  return result;
}

// Hasil akhir: teks asli + terjemahan miring dalam tanda kurung
async function buildCombined(original, position) {
  const raw = await translateCached(original);
  const bracketed = `(${toUnicodeItalic(raw)})`;
  const sep = position === "newline" ? "\n" : " ";
  return `${original.trim()}${sep}${bracketed}`;
}

function supported(version) {
  return Office.context.requirements.isSetSupported("ExcelApi", version);
}

// ---------- Sel ----------
async function translateRange(context, range, position, stats) {
  range.load("values,formulas,rowCount,columnCount");
  await context.sync();

  for (let r = 0; r < range.rowCount; r++) {
    for (let c = 0; c < range.columnCount; c++) {
      const value = range.values[r][c];
      const formula = range.formulas[r][c];
      // Lewati sel berisi rumus agar rumus tidak tertimpa nilai statis
      if (typeof formula === "string" && formula.startsWith("=")) continue;
      if (!isTranslatable(value)) continue;

      const combined = await buildCombined(value, position);
      const cell = range.getCell(r, c);
      cell.values = [[combined]];
      if (position === "newline") cell.format.wrapText = true;
      stats.cells++;
    }
  }
  await context.sync();
}

// ---------- Shape (kotak teks, bentuk, grup) ----------
async function translateShapeCollection(context, shapes, position, stats) {
  shapes.load("items/name,items/type");
  await context.sync();

  for (const shape of shapes.items) {
    try {
      if (shape.type === "Group") {
        await translateShapeCollection(context, shape.group.shapes, position, stats);
        continue;
      }
      if (shape.type !== "GeometricShape") continue; // garis/gambar tidak punya teks

      const textRange = shape.textFrame.textRange;
      textRange.load("text");
      await context.sync();

      if (!isTranslatable(textRange.text)) continue;
      textRange.text = await buildCombined(textRange.text, position);
      await context.sync();
      stats.shapes++;
    } catch (e) {
      console.warn("Shape dilewati:", e);
    }
  }
}

// ---------- Grafik (judul + judul sumbu) ----------
async function translateTitle(context, title, position) {
  title.load("text,visible");
  await context.sync();
  if (!title.visible || !isTranslatable(title.text)) return false;
  title.text = await buildCombined(title.text, position);
  await context.sync();
  return true;
}

async function translateCharts(context, charts, position, stats) {
  charts.load("items/name");
  await context.sync();

  for (const chart of charts.items) {
    const titles = [chart.title];
    try {
      titles.push(chart.axes.categoryAxis.title, chart.axes.valueAxis.title);
    } catch (e) { /* grafik tanpa sumbu (mis. pie) */ }

    for (const title of titles) {
      try {
        if (await translateTitle(context, title, position)) stats.charts++;
      } catch (e) {
        console.warn("Judul grafik dilewati:", e);
      }
    }
  }
}

// ---------- Komentar & catatan ----------
async function translateComments(context, sheets, position, stats) {
  if (supported("1.10")) {
    try {
      const comments = context.workbook.comments;
      comments.load("items/content");
      await context.sync();
      for (const comment of comments.items) {
        if (!isTranslatable(comment.content)) continue;
        comment.content = await buildCombined(comment.content, position);
        stats.comments++;
      }
      await context.sync();
    } catch (e) {
      console.warn("Komentar dilewati:", e);
    }
  }

  if (supported("1.18")) {
    for (const ws of sheets) {
      try {
        const notes = ws.notes;
        notes.load("items/content");
        await context.sync();
        for (const note of notes.items) {
          if (!isTranslatable(note.content)) continue;
          note.content = await buildCombined(note.content, position);
          stats.comments++;
        }
        await context.sync();
      } catch (e) {
        console.warn("Catatan dilewati:", e);
      }
    }
  }
}

async function handleTranslate() {
  const statusDiv = document.getElementById("status");
  statusDiv.className = "";
  statusDiv.innerText = "Menerjemahkan...";

  const position = document.querySelector('input[name="position"]:checked').value;
  const scope = document.querySelector('input[name="scope"]:checked').value;
  const want = {
    cells: document.getElementById("optCells").checked,
    shapes: document.getElementById("optShapes").checked,
    charts: document.getElementById("optCharts").checked,
    comments: document.getElementById("optComments").checked,
  };

  const stats = { cells: 0, shapes: 0, charts: 0, comments: 0 };

  try {
    await Excel.run(async (context) => {
      // Tentukan worksheet yang diproses
      let sheets;
      if (scope === "workbook") {
        context.workbook.worksheets.load("items");
        await context.sync();
        sheets = context.workbook.worksheets.items;
      } else {
        sheets = [context.workbook.worksheets.getActiveWorksheet()];
      }

      // Mode "Sel terpilih": hanya sel pada seleksi
      if (scope === "selection") {
        if (want.cells) {
          await translateRange(context, context.workbook.getSelectedRange(), position, stats);
        }
      } else {
        for (const ws of sheets) {
          if (want.cells) {
            const used = ws.getUsedRangeOrNullObject();
            used.load("isNullObject");
            await context.sync();
            if (!used.isNullObject) await translateRange(context, used, position, stats);
          }
          if (want.shapes && supported("1.9")) {
            await translateShapeCollection(context, ws.shapes, position, stats);
          }
          if (want.charts) {
            await translateCharts(context, ws.charts, position, stats);
          }
        }
        if (want.comments) {
          await translateComments(context, sheets, position, stats);
        }
      }

      await context.sync();
    });

    const parts = [];
    if (stats.cells) parts.push(`${stats.cells} sel`);
    if (stats.shapes) parts.push(`${stats.shapes} shape`);
    if (stats.charts) parts.push(`${stats.charts} judul grafik`);
    if (stats.comments) parts.push(`${stats.comments} komentar/catatan`);

    statusDiv.className = "success";
    statusDiv.innerText = parts.length
      ? `✓ Selesai: ${parts.join(", ")}.`
      : "Tidak ada teks yang perlu diterjemahkan.";
  } catch (error) {
    console.error(error);
    statusDiv.className = "error";
    statusDiv.innerText = "Terjadi kendala: " + error.message;
  }
}
