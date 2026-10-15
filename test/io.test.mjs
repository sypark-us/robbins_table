import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseCSV, toCSV, sniffDelimiter } from '../js/csv.js';
import { parseXml, child, children, attr, textOf, x } from '../js/xml.js';
import { zipWrite, zipRead, crc32, utf8 } from '../js/zip.js';
import { exportXlsx, importXlsx } from '../js/xlsx.js';
import { Workbook } from '../js/model.js';
import { dateToSerial } from '../js/format.js';

test('csv parse and serialise', () => {
  const rows = parseCSV('a,b,c\r\n1,"x, y","he said ""hi"""\n2,,3\n');
  assert.deepEqual(rows, [['a', 'b', 'c'], ['1', 'x, y', 'he said "hi"'], ['2', '', '3']]);
  assert.equal(sniffDelimiter('a;b;c\n1;2;3'), ';');
  assert.equal(sniffDelimiter('a\tb\n1\t2'), '\t');
  const csv = toCSV([['a', 'b,c'], [1, 'q"q'], [null, 'multi\nline']]);
  assert.equal(csv, 'a,"b,c"\r\n1,"q""q"\r\n,"multi\nline"\r\n');
  assert.deepEqual(parseCSV(csv), [['a', 'b,c'], ['1', 'q"q'], ['', 'multi\nline']]);
  assert.deepEqual(parseCSV('﻿x,y'), [['x', 'y']]);
  assert.deepEqual(parseCSV('"quoted\nnewline",2'), [['quoted\nnewline', '2']]);
});

test('xml parser', () => {
  const doc = parseXml('<?xml version="1.0"?><root a="1" b=\'two\'><!-- c --><item x="&amp;&lt;">text &#65;<sub/></item><item>2</item></root>');
  assert.equal(doc.name, 'root'); assert.equal(attr(doc, 'a'), '1'); assert.equal(attr(doc, 'b'), 'two');
  const items = children(doc, 'item');
  assert.equal(items.length, 2); assert.equal(attr(items[0], 'x'), '&<'); assert.equal(textOf(items[0]), 'text A');
  assert.equal(child(items[0], 'sub').name, 'sub');
  const ns = parseXml('<x:a xmlns:x="u"><x:b r:id="rId1" xmlns:r="v">t</x:b></x:a>');
  assert.equal(attr(child(ns, 'b'), 'id'), 'rId1');
  assert.equal(x('c', { r: 'A1', t: 'str' }, [x('v', {}, 'a&lt;b')]), '<c r="A1" t="str"><v>a&lt;b</v></c>');
  assert.equal(x('c', { r: 'a"b' }), '<c r="a&quot;b"/>');
  assert.equal(x('empty', { a: undefined }), '<empty/>');
});

test('zip round trip', async () => {
  assert.equal(crc32(utf8.encode('hello')), 0x3610a686);
  const big = 'The quick brown fox jumps over the lazy dog. '.repeat(200);
  const z = await zipWrite([{ name: 'a.txt', data: 'hello' }, { name: 'dir/b.txt', data: big }, { name: 'bin', data: new Uint8Array([1, 2, 3]) }]);
  assert.equal(z[0], 0x50); assert.equal(z[1], 0x4b);
  const back = await zipRead(z);
  assert.equal(utf8.decode(back.get('a.txt')), 'hello');
  assert.equal(utf8.decode(back.get('dir/b.txt')), big);
  assert.deepEqual([...back.get('bin')], [1, 2, 3]);
  assert.ok(z.length < big.length, 'deflate should compress');
});

test('xlsx export/import round trip', async () => {
  const wb = Workbook.blank(2);
  const s = wb.sheets[0];
  wb.transact('fill', () => {
    s.setInput(0, 0, 'Name'); s.setInput(0, 1, 'Qty'); s.setInput(0, 2, 'Price'); s.setInput(0, 3, 'Total');
    s.setInput(1, 0, 'Widget'); s.setInput(1, 1, '3'); s.setInput(1, 2, '2.5'); s.setInput(1, 3, '=B2*C2');
    s.setInput(2, 0, 'Gadget "X"'); s.setInput(2, 1, '2'); s.setInput(2, 2, '10'); s.setInput(2, 3, '=B3*C3');
    s.setInput(3, 3, '=SUM(D2:D3)'); s.setInput(4, 0, 'TRUE'); s.setInput(4, 1, '2024-03-15'); s.setInput(4, 2, '50%');
    s.setInput(5, 0, '=1/0'); s.setInput(5, 1, '=Sheet2!A1*2'); s.setInput(6, 0, 'ünïcödé <&>');
    s.applyStyle({ r1: 0, c1: 0, r2: 0, c2: 3 }, { bold: true, fill: '#ddebf7', align: 'center', border: { bottom: { style: 'thin', color: '#000000' } } });
    s.applyStyle({ r1: 1, c1: 2, r2: 3, c2: 3 }, { numFmt: '$#,##0.00' });
    s.applyStyle({ r1: 1, c1: 0, r2: 1, c2: 0 }, { italic: true, color: '#c00000', size: 14, font: 'Arial', wrap: true, valign: 'top' });
    s.setColWidth(0, 150); s.setRowHeight(0, 30); s.setColHidden(5, true);
    s.merge({ r1: 7, c1: 0, r2: 8, c2: 2 });
    s.setFreeze(1, 1);
    s.setCF([{ id: 'x', r1: 1, c1: 1, r2: 3, c2: 1, type: 'cellIs', op: 'gt', v1: '2', style: { fill: '#c6efce', color: '#006100', bold: true } }, { id: 'y', r1: 1, c1: 2, r2: 3, c2: 2, type: 'colorScale', colors: ['#ff0000', '#00ff00'] }]);
    s.setValidations([{ id: 'v', r1: 1, c1: 0, r2: 3, c2: 0, type: 'list', v1: 'Widget,Gadget', allowBlank: true, showDropdown: true, errorStyle: 'stop' }]);
    s.setFilter({ r1: 0, c1: 0, r2: 3, c2: 3, conds: {} });
    s.setHyperlink(6, 1, 'https://example.com');
    wb.sheets[1].setInput(0, 0, '21');
    wb.setName('Total', '=Sheet1!$D$4');
    wb.renameSheet(wb.sheets[1], 'Data 2');
  });
  const bytes = await exportXlsx(wb);
  const wb2 = await importXlsx(bytes, 'test.xlsx');
  const t = wb2.sheets[0];
  assert.equal(wb2.sheets.length, 2); assert.equal(wb2.sheets[1].name, 'Data 2');
  assert.equal(t.getValue(0, 0), 'Name'); assert.equal(t.getValue(1, 0), 'Widget'); assert.equal(t.getValue(2, 0), 'Gadget "X"');
  assert.equal(t.getValue(1, 1), 3); assert.equal(t.getCell(1, 3).f, '=B2*C2'); assert.equal(t.getValue(1, 3), 7.5);
  assert.equal(t.getValue(3, 3), 27.5); assert.equal(t.getValue(4, 0), true); assert.equal(t.getValue(4, 1), dateToSerial(2024, 3, 15));
  assert.equal(t.getStyle(4, 1).numFmt, 'yyyy-mm-dd'); assert.equal(t.getValue(4, 2), 0.5); assert.equal(t.getStyle(4, 2).numFmt, '0%');
  assert.equal(t.getValue(5, 0).error, '#DIV/0!'); assert.equal(t.getCell(5, 1).f, "='Data 2'!A1*2"); assert.equal(t.getValue(5, 1), 42);
  assert.equal(t.getValue(6, 0), 'ünïcödé <&>');
  const h = t.getStyle(0, 0);
  assert.equal(h.bold, true); assert.equal(h.fill, '#ddebf7'); assert.equal(h.align, 'center'); assert.equal(h.border.bottom.style, 'thin');
  assert.equal(t.getStyle(1, 2).numFmt, '$#,##0.00');
  const i = t.getStyle(1, 0); assert.equal(i.italic, true); assert.equal(i.color, '#c00000'); assert.equal(i.size, 14); assert.equal(i.font, 'Arial'); assert.equal(i.wrap, true); assert.equal(i.valign, 'top');
  assert.equal(t.colWidth(0), 150); assert.equal(t.rowHeight(0), 30); assert.equal(t.colAxis.hidden.has(5), true);
  assert.deepEqual(t.merges, [{ r1: 7, c1: 0, r2: 8, c2: 2 }]);
  assert.deepEqual(t.freeze, { rows: 1, cols: 1 });
  assert.equal(t.cf.length, 2); assert.equal(t.cf[0].type, 'cellIs'); assert.equal(t.cf[0].op, 'gt'); assert.equal(t.cf[0].v1, '=2'); assert.equal(t.cf[0].style.fill, '#c6efce'); assert.equal(t.cf[0].style.bold, true);
  assert.equal(t.cf[1].type, 'colorScale'); assert.deepEqual(t.cf[1].colors, ['#ff0000', '#00ff00']);
  assert.equal(t.validations.length, 1); assert.equal(t.validations[0].type, 'list'); assert.equal(t.validations[0].v1, 'Widget,Gadget');
  assert.deepEqual({ r1: t.filter.r1, c2: t.filter.c2 }, { r1: 0, c2: 3 });
  assert.equal(t.getCell(6, 1).hyperlink, 'https://example.com');
  assert.equal(wb2.names.get('TOTAL').ref, '=Sheet1!$D$4');
  // shared strings / shared formulas / theme colors from a hand-built minimal xlsx
});

test('xlsx import handles shared strings, shared formulas and styles from other producers', async () => {
  const { zipWrite: zw } = await import('../js/zip.js');
  const files = [
    { name: '[Content_Types].xml', data: '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>' },
    { name: 'xl/workbook.xml', data: '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>' },
    { name: 'xl/_rels/workbook.xml.rels', data: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="w" Target="worksheets/sheet1.xml"/></Relationships>' },
    { name: 'xl/sharedStrings.xml', data: '<sst><si><t>Hello</t></si><si><r><t>Ri</t></r><r><t>ch</t></r></si></sst>' },
    { name: 'xl/styles.xml', data: '<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts><fonts><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="12"/><color theme="4"/><name val="Arial"/></font></fonts><fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor indexed="13"/></patternFill></fill></fills><borders><border/></borders><cellXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="164" fontId="1" fillId="2" borderId="0" applyNumberFormat="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>' },
    { name: 'xl/worksheets/sheet1.xml', data: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" s="1"><v>1.5</v></c><c r="D1" s="2"><v>45000</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">A2*10</f><v>10</v></c></row><row r="3"><c r="A3"><v>2</v></c><c r="B3"><f t="shared" si="0"/><v>20</v></c></row></sheetData></worksheet>' },
  ];
  const wb = await importXlsx(await zw(files));
  const s = wb.sheets[0];
  assert.equal(s.getValue(0, 0), 'Hello'); assert.equal(s.getValue(0, 1), 'Rich');
  assert.equal(s.getStyle(0, 2).numFmt, '0.000'); assert.equal(s.getStyle(0, 2).bold, true); assert.equal(s.getStyle(0, 2).font, 'Arial'); assert.equal(s.getStyle(0, 2).size, 12); assert.equal(s.getStyle(0, 2).color, '#4472c4'); assert.equal(s.getStyle(0, 2).fill, '#ffff00');
  assert.equal(s.getStyle(0, 3).numFmt, 'm/d/yyyy');
  assert.equal(s.getCell(2, 1).f, '=A3*10'); assert.equal(s.getValue(2, 1), 20);
});
