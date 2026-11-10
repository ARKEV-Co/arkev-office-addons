// The commands ARKEV's plugin runs inside the editor with Asc.plugin.callCommand. Each function is
// sent to the editor as source text and runs there with only the Document Builder API (Api) and
// Asc.scope, so each one is self-contained: no closures, no helpers from outside it.
// The outline and the edit list are the "Office edit spec v1" (arkev-core docs/office-editors.md);
// the arkev-docs worker gives the same outline and applies the same edits when no editor is open.
// Every command returns JSON text.
(function (window) {
  "use strict";

  // ---------------------------------------------------------------- Word (docx)

  function wordOutline() {
    var MAXB = 400, MAXT = 2000, TOTAL = 300000;
    var doc = Api.GetDocument(), n = doc.GetElementsCount(), blocks = [], total = 0, truncated = false;
    var arabic = /[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]/;
    function cut(s) { s = s == null ? "" : String(s); total += s.length; return s.length > MAXT ? s.slice(0, MAXT) : s; }
    function contentText(content) {
      var parts = [], k, m = content.GetElementsCount();
      for (k = 0; k < m; k++) {
        var e = content.GetElement(k);
        if (e && e.GetText) parts.push(e.GetText({Numbering: false}).replace(/\r?\n$/, ""));
      }
      return parts.join("\n");
    }
    for (var i = 0; i < n; i++) {
      if (i >= MAXB || total > TOTAL) { truncated = true; break; }
      var el = doc.GetElement(i), type = el.GetClassType();
      if (type === "table") {
        var rows = [], rc = el.GetRowsCount();
        for (var r = 0; r < rc && r < 200; r++) {
          var row = el.GetRow(r), cc = row.GetCellsCount(), cells = [];
          for (var c = 0; c < cc && c < 30; c++) cells.push(cut(contentText(row.GetCell(c).GetContent())));
          rows.push(cells);
        }
        blocks.push({i: i, type: "table", rows: rows});
      } else {
        var text = "", style = "";
        if (type === "paragraph") {
          text = el.GetText({Numbering: false}).replace(/\r?\n$/, "");
          try { var st = el.GetStyle(); style = st ? st.GetName() : ""; } catch (e) { style = ""; }
        } else if (el.GetContent) {
          text = contentText(el.GetContent());
        }
        blocks.push({i: i, type: "paragraph", style: style, text: cut(text), rtl: arabic.test(text)});
      }
    }
    var selection = {text: ""};
    try { var range = doc.GetRangeBySelect(); if (range) selection.text = cut(range.GetText()); } catch (e) {}
    return JSON.stringify({format: "docx", blocks: blocks, truncated: truncated, selection: selection});
  }

  function wordApply() {
    var edits = Asc.scope.edits || [], doc = Api.GetDocument(), applied = 0, skipped = [];
    var arabic = /[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]/;
    var STYLES = {title: "Title", heading1: "Heading 1", heading2: "Heading 2", heading3: "Heading 3", normal: "Normal"};
    function skip(i, why) { skipped.push({index: i, reason: why}); }
    function block(i, b) {
      if (typeof b !== "number" || b < 0 || b >= doc.GetElementsCount() || Math.floor(b) !== b) throw "block " + b + " is out of range (the document has " + doc.GetElementsCount() + " blocks)";
      return doc.GetElement(b);
    }
    function paragraph(i, b) {
      var el = block(i, b);
      if (el.GetClassType() !== "paragraph") throw "block " + b + " is not a paragraph";
      return el;
    }
    // Text with Arabic-script letters (Sorani, Arabic): a right-to-left paragraph in Noto Sans
    // Arabic, as the arkev-docs worker writes it. In a right-to-left paragraph Word reads the
    // alignment logically, so a visual "right" is stored as "left" (and the default start is right).
    function rtl(p, text, align) {
      var r = arabic.test(String(text));
      if (r) {
        try { p.SetReadingOrder("rtl"); } catch (e) {}
        try {
          for (var k = 0; k < p.GetElementsCount(); k++) {
            var el = p.GetElement(k);
            if (el && el.SetFontFamily) el.SetFontFamily("Noto Sans Arabic");
          }
        } catch (e) {}
      }
      if (align) p.SetJc(r ? ({left: "right", right: "left"}[align] || align) : align);
    }
    function setText(p, text) {
      var first = p.GetElementsCount() > 0 ? p.GetElement(0) : null, pr = null;
      try { if (first && first.GetTextPr) pr = first.GetTextPr(); } catch (e) { pr = null; }
      p.RemoveAllElements();
      var run = Api.CreateRun();
      run.AddText(String(text));
      if (pr) { try { run.SetTextPr(pr); } catch (e) {} }
      p.AddElement(run);
    }
    function place(e, el) {
      var pos = e.position || "end";
      if (pos === "end") { doc.Push(el); return; }
      if (pos === "start") { doc.AddElement(0, el); return; }
      block(-1, e.block);
      doc.AddElement(pos === "before" ? e.block : e.block + 1, el);
    }
    function newParagraph(e) {
      var p = Api.CreateParagraph(), style = e.style || "normal";
      p.AddText(String(e.text));
      if (style === "bullet" || style === "numbered") {
        var num = doc.CreateNumbering(style === "bullet" ? "bullet" : "numbered");
        p.SetNumbering(num.GetLevel(0));
      } else if (STYLES[style]) {
        var st = doc.GetStyle(STYLES[style]);
        if (st) p.SetStyle(st);
      }
      rtl(p, e.text, "");
      return p;
    }
    function color(hex) { return [parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16)]; }
    for (var i = 0; i < edits.length; i++) {
      var e = edits[i];
      try {
        switch (e.op) {
          case "replace_text": {
            var found = doc.Search(String(e.find), e.match_case !== false);
            if (!found || !found.length) throw "\"" + String(e.find).slice(0, 60) + "\" was not found";
            doc.SearchAndReplace({searchString: String(e.find), replaceString: String(e.replace), matchCase: e.match_case !== false});
            break;
          }
          case "insert_paragraph":
            place(e, newParagraph(e));
            break;
          case "replace_block": {
            var p = paragraph(i, e.block);
            setText(p, e.text);
            rtl(p, e.text, "");
            break;
          }
          case "delete_block":
            block(i, e.block);
            doc.RemoveElement(e.block);
            break;
          case "set_table_cell": {
            var t = block(i, e.block);
            if (t.GetClassType() !== "table") throw "block " + e.block + " is not a table";
            if (e.row < 0 || e.row >= t.GetRowsCount() || e.col < 0 || e.col >= t.GetRow(e.row).GetCellsCount()) throw "cell " + e.row + "," + e.col + " is outside the table";
            var content = t.GetCell(e.row, e.col).GetContent();
            while (content.GetElementsCount() > 1) content.RemoveElement(1);
            var cp = content.GetElement(0);
            if (!cp || cp.GetClassType() !== "paragraph") { content.RemoveAllElements(); cp = Api.CreateParagraph(); content.Push(cp); }
            setText(cp, e.text);
            rtl(cp, e.text, "");
            break;
          }
          case "insert_table": {
            var rows = e.rows || [], cols = 0, r, c;
            for (r = 0; r < rows.length; r++) cols = Math.max(cols, rows[r].length);
            if (!rows.length || !cols) throw "the table has no cells";
            var table = Api.CreateTable(cols, rows.length);
            try { var grid = doc.GetStyle("Table Grid"); if (grid) table.SetStyle(grid); } catch (x) {}
            table.SetWidth("percent", 100);
            for (r = 0; r < rows.length; r++) {
              for (c = 0; c < cols; c++) {
                var cell = table.GetCell(r, c), tp = cell.GetContent().GetElement(0), v = rows[r][c] == null ? "" : String(rows[r][c]);
                tp.AddText(v);
                if (r === 0 && e.header) tp.SetBold(true);
                rtl(tp, v, "");
              }
            }
            place(e, table);
            break;
          }
          case "format_block": {
            var fp = paragraph(i, e.block), range = fp.GetRange();
            if (typeof e.bold === "boolean") range.SetBold(e.bold);
            if (typeof e.italic === "boolean") range.SetItalic(e.italic);
            if (typeof e.underline === "boolean") range.SetUnderline(e.underline);
            if (typeof e.size_pt === "number") range.SetFontSize(Math.round(e.size_pt * 2));
            if (e.color) { var rgb = color(e.color); range.SetColor(rgb[0], rgb[1], rgb[2]); }
            if (e.align) {
              var a = e.align === "justify" ? "both" : e.align;
              if (arabic.test(fp.GetText())) a = {left: "right", right: "left"}[a] || a;
              fp.SetJc(a);
            }
            break;
          }
          default:
            throw "unknown op " + e.op;
        }
        applied++;
      } catch (err) {
        skip(i, typeof err === "string" ? err : String(err && err.message || err));
      }
    }
    return JSON.stringify({applied: applied, skipped: skipped});
  }

  // ---------------------------------------------------------------- Excel (xlsx)

  function cellOutline() {
    var MAXR = 200, MAXC = 30, MAXT = 2000;
    var sheets = Api.GetSheets(), active = Api.GetActiveSheet().GetName(), out = [], truncated = sheets.length > 10;
    function letters(n) { var s = ""; n++; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
    for (var s = 0; s < sheets.length && s < 10; s++) {
      var sh = sheets[s], used = sh.GetUsedRange(), rows = [], formulas = {}, has = false;
      var r0 = used.GetRow() - 1, c0 = used.GetCol() - 1;
      var nr = used.GetRowsCount ? used.GetRowsCount() : used.GetRows().GetCount ? used.GetRows().GetCount() : 1;
      var nc = used.GetColumnsCount ? used.GetColumnsCount() : used.GetCols().GetCount ? used.GetCols().GetCount() : 1;
      if (nr > MAXR || nc > MAXC) truncated = true;
      for (var r = 0; r < nr && r < MAXR; r++) {
        var row = [];
        for (var c = 0; c < nc && c < MAXC; c++) {
          var cell = sh.GetRangeByNumber(r0 + r, c0 + c), v = cell.GetValue2 ? cell.GetValue2() : cell.GetValue(), f = cell.GetFormula();
          if (typeof f === "string" && f.charAt(0) === "=") { formulas[letters(c0 + c) + (r0 + r + 1)] = f; has = true; }
          if (typeof v === "string" && v.length > MAXT) v = v.slice(0, MAXT);
          row.push(v === "" ? null : v);
        }
        rows.push(row);
      }
      var item = {name: sh.GetName(), active: sh.GetName() === active, used: used.GetAddress(false, false, "xlA1", false), rows: rows};
      if (has) item.formulas = formulas;
      out.push(item);
    }
    var selection = {sheet: active, range: ""};
    try { selection.range = Api.GetSelection().GetAddress(false, false, "xlA1", false); } catch (e) {}
    return JSON.stringify({format: "xlsx", sheets: out, truncated: truncated, selection: selection});
  }

  function cellApply() {
    var edits = Asc.scope.edits || [], applied = 0, skipped = [];
    var arabic = /[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]/;
    function sheet(name) {
      var sh = name ? Api.GetSheet(String(name)) : Api.GetActiveSheet();
      if (!sh) throw "there is no sheet \"" + name + "\"";
      return sh;
    }
    function rgb(hex) { return Api.CreateColorFromRGB(parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16)); }
    for (var i = 0; i < edits.length; i++) {
      var e = edits[i];
      try {
        switch (e.op) {
          case "set_cells": {
            var sh = sheet(e.sheet), start = sh.GetRange(String(e.cell));
            if (!start) throw "\"" + e.cell + "\" is not a cell";
            var r0 = start.GetRow() - 1, c0 = start.GetCol() - 1, vals = e.values || [];
            for (var r = 0; r < vals.length; r++) {
              for (var c = 0; c < vals[r].length; c++) {
                var cell = sh.GetRangeByNumber(r0 + r, c0 + c), v = vals[r][c];
                if (v === null || v === undefined) cell.Clear();
                else {
                  cell.SetValue(v);
                  if (typeof v === "string" && arabic.test(v)) { try { cell.SetAlignHorizontal("right"); } catch (x) {} }
                }
              }
            }
            break;
          }
          case "add_sheet":
            if (Api.GetSheet(String(e.name))) throw "a sheet \"" + e.name + "\" already exists";
            Api.AddSheet(String(e.name));
            break;
          case "format_range": {
            var rg = sheet(e.sheet).GetRange(String(e.range));
            if (!rg) throw "\"" + e.range + "\" is not a range";
            if (typeof e.bold === "boolean") rg.SetBold(e.bold);
            if (typeof e.italic === "boolean") rg.SetItalic(e.italic);
            if (e.fill) rg.SetFillColor(rgb(e.fill));
            if (e.color) rg.SetFontColor(rgb(e.color));
            if (e.number_format) rg.SetNumberFormat(String(e.number_format));
            if (e.align) rg.SetAlignHorizontal(e.align);
            break;
          }
          default:
            throw "unknown op " + e.op;
        }
        applied++;
      } catch (err) {
        skipped.push({index: i, reason: typeof err === "string" ? err : String(err && err.message || err)});
      }
    }
    return JSON.stringify({applied: applied, skipped: skipped});
  }

  // ---------------------------------------------------------------- PowerPoint (pptx)

  function slideOutline() {
    var MAXS = 60, MAXT = 2000;
    var pres = Api.GetPresentation(), n = pres.GetSlidesCount(), slides = [];
    function isTitle(sh) {
      try { var ph = sh.GetPlaceholder && sh.GetPlaceholder(); var t = ph ? ph.GetType() : ""; return t === "title" || t === "ctrTitle"; } catch (e) { return false; }
    }
    function paragraphs(sh) {
      var out = [];
      try {
        var content = sh.GetDocContent && sh.GetDocContent();
        if (!content) return out;
        for (var k = 0; k < content.GetElementsCount(); k++) {
          var p = content.GetElement(k);
          var t = p && p.GetText ? p.GetText({Numbering: false}).replace(/\r?\n$/, "") : "";
          if (t) out.push(t.length > MAXT ? t.slice(0, MAXT) : t);
        }
      } catch (e) {}
      return out;
    }
    for (var i = 0; i < n && i < MAXS; i++) {
      var slide = pres.GetSlideByIndex(i), shapes = slide.GetAllShapes(), title = "", texts = [];
      for (var s = 0; s < shapes.length; s++) {
        var ps = paragraphs(shapes[s]);
        if (!title && isTitle(shapes[s])) title = ps.join("\n");
        else texts = texts.concat(ps);
      }
      slides.push({i: i, title: title, texts: texts});
    }
    var selection = {slide: 0};
    try { selection.slide = pres.GetCurSlideIndex ? pres.GetCurSlideIndex() : 0; } catch (e) {}
    return JSON.stringify({format: "pptx", slides: slides, truncated: n > MAXS, selection: selection});
  }

  function slideApply() {
    var edits = Asc.scope.edits || [], pres = Api.GetPresentation(), applied = 0, skipped = [];
    var arabic = /[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]/;
    function slideAt(i) {
      var n = pres.GetSlidesCount();
      if (typeof i !== "number" || i < 0 || i >= n || Math.floor(i) !== i) throw "slide " + i + " is out of range (the deck has " + n + " slides)";
      return pres.GetSlideByIndex(i);
    }
    function phType(sh) { try { var ph = sh.GetPlaceholder && sh.GetPlaceholder(); return ph ? ph.GetType() : ""; } catch (e) { return ""; } }
    function fill(sh, lines) {
      var content = sh.GetDocContent();
      content.RemoveAllElements();
      for (var k = 0; k < lines.length; k++) {
        var p = Api.CreateParagraph();
        p.AddText(String(lines[k]));
        if (arabic.test(String(lines[k]))) { try { p.SetJc("right"); } catch (x) {} }
        content.Push(p);
      }
    }
    // The title placeholder and the body: a body, object (content) or subtitle placeholder, or
    // else the first other placeholder that holds text (never the date, footer or slide number).
    function parts(slide) {
      var shapes = slide.GetAllShapes(), title = null, body = null, other = null;
      var NOT_BODY = {date: true, footer: true, header: true, sldNumber: true, sldImage: true, title: true, ctrTitle: true, "": true};
      for (var s = 0; s < shapes.length; s++) {
        var t = phType(shapes[s]);
        if (!title && (t === "title" || t === "ctrTitle")) title = shapes[s];
        else if (!body && (t === "body" || t === "object" || t === "subTitle")) body = shapes[s];
        else if (!other && !NOT_BODY[t] && shapes[s].GetDocContent && shapes[s].GetDocContent()) other = shapes[s];
      }
      return {title: title, body: body || other};
    }
    function layout() {
      try {
        var master = pres.GetMaster(0);
        for (var l = 0; l < master.GetLayoutsCount(); l++) {
          var lay = master.GetLayout(l);
          var name = (lay.GetName && lay.GetName()) || "";
          if (/title.*content|title, content|obj/i.test(name)) return lay;
        }
        return master.GetLayoutsCount() > 1 ? master.GetLayout(1) : null;
      } catch (e) { return null; }
    }
    for (var i = 0; i < edits.length; i++) {
      var e = edits[i];
      try {
        switch (e.op) {
          case "add_slide": {
            var slide = Api.CreateSlide(), lay = layout();
            if (lay && slide.ApplyLayout) slide.ApplyLayout(lay);
            var at = typeof e.after === "number" ? e.after + 1 : pres.GetSlidesCount();
            if (at < 0 || at > pres.GetSlidesCount()) throw "slide " + e.after + " is out of range";
            pres.AddSlide(slide);
            if (at !== pres.GetSlidesCount() - 1 && slide.MoveTo) slide.MoveTo(at);
            var ps = parts(slide), W = pres.GetWidth ? pres.GetWidth() : 9144000, H = pres.GetHeight ? pres.GetHeight() : 6858000;
            if (!ps.title) {
              ps.title = Api.CreateShape("rect", W * 0.9, H * 0.15, Api.CreateNoFill(), Api.CreateStroke(0, Api.CreateNoFill()));
              ps.title.SetPosition(W * 0.05, H * 0.05);
              slide.AddObject(ps.title);
            }
            if (!ps.body) {
              ps.body = Api.CreateShape("rect", W * 0.9, H * 0.65, Api.CreateNoFill(), Api.CreateStroke(0, Api.CreateNoFill()));
              ps.body.SetPosition(W * 0.05, H * 0.25);
              slide.AddObject(ps.body);
            }
            fill(ps.title, [e.title || ""]);
            fill(ps.body, e.bullets || []);
            break;
          }
          case "set_slide_text": {
            var p2 = parts(slideAt(e.slide));
            if (typeof e.title === "string" && !p2.title) throw "slide " + e.slide + " has no title placeholder";
            if (e.bullets && !p2.body) throw "slide " + e.slide + " has no body placeholder";
            if (typeof e.title === "string") fill(p2.title, [e.title]);
            if (e.bullets) fill(p2.body, e.bullets);
            break;
          }
          case "delete_slide":
            slideAt(e.slide).Delete();
            break;
          default:
            throw "unknown op " + e.op;
        }
        applied++;
      } catch (err) {
        skipped.push({index: i, reason: typeof err === "string" ? err : String(err && err.message || err)});
      }
    }
    return JSON.stringify({applied: applied, skipped: skipped});
  }

  window.ARKEV_COMMANDS = Object.freeze({
    word: {outline: wordOutline, apply: wordApply},
    cell: {outline: cellOutline, apply: cellApply},
    slide: {outline: slideOutline, apply: slideApply}
  });
})(window);
