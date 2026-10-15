import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../js/model.js';
import { parse, tokenize, shiftFormula, adjustFormulaForInsertDelete, extractRefs } from '../js/formula.js';
import { dateToSerial } from '../js/format.js';

function book() { return Workbook.blank(2); }
function set(sh, a1, text) { const { r, c } = ref(a1); sh.wb.transact('set', () => sh.setInput(r, c, text)); }
function ref(a1) { const m = /^([A-Z]+)(\d+)$/.exec(a1); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return { r: +m[2] - 1, c: c - 1 }; }
function val(sh, a1) { const { r, c } = ref(a1); const v = sh.getValue(r, c); return v && v.error ? v.error : v; }
function f(a1, formula) { const wb = book(); const sh = wb.sheets[0]; set(sh, a1, formula); return val(sh, a1); }
const approx = (a, b, eps = 1e-9) => assert.ok(Math.abs(a - b) < eps, `${a} != ${b}`);

test('tokenizer and parser', () => {
  const t = tokenize('SUM(A1:B2, $C$3)*2');
  assert.equal(t.filter((x) => x.t === 'ref').length, 2);
  const ast = parse('=1+2*3');
  assert.equal(ast.type, 'bin'); assert.equal(ast.op, '+');
  assert.equal(parse("='My Sheet'!A1").sheet, 'My Sheet');
  assert.deepEqual(extractRefs('=A1+Sheet2!B2:C3').map((r) => r.text), ['A1', 'Sheet2!B2:C3']);
  assert.throws(() => parse('=1+'));
  assert.equal(parse('=A:A').kind, 'col');
  assert.equal(parse('=3:5').kind, 'row');
});

test('arithmetic, precedence and operators', () => {
  assert.equal(f('A1', '=1+2*3'), 7);
  assert.equal(f('A1', '=(1+2)*3'), 9);
  assert.equal(f('A1', '=-2^2'), 4);
  assert.equal(f('A1', '=2^3^2'), 64);
  assert.equal(f('A1', '=10/4'), 2.5);
  assert.equal(f('A1', '=1/0'), '#DIV/0!');
  assert.equal(f('A1', '="a"&"b"&1'), 'ab1');
  assert.equal(f('A1', '=50%'), 0.5);
  assert.equal(f('A1', '="3"+4'), 7);
  assert.equal(f('A1', '="x"+4'), '#VALUE!');
  assert.equal(f('A1', '=TRUE+1'), 2);
  assert.equal(f('A1', '=1=1'), true);
  assert.equal(f('A1', '="a"<"B"'), true);
  assert.equal(f('A1', '="A"="a"'), true);
  assert.equal(f('A1', '=1<>2'), true);
  assert.equal(f('A1', '=1<"a"'), true);
});

test('cell references and recalculation', () => {
  const wb = book(); const sh = wb.sheets[0];
  set(sh, 'A1', '10'); set(sh, 'A2', '20'); set(sh, 'A3', '=A1+A2'); set(sh, 'A4', '=A3*2');
  assert.equal(val(sh, 'A3'), 30); assert.equal(val(sh, 'A4'), 60);
  set(sh, 'A1', '5');
  assert.equal(val(sh, 'A3'), 25); assert.equal(val(sh, 'A4'), 50);
  set(sh, 'B1', '=SUM(A1:A2)'); assert.equal(val(sh, 'B1'), 25);
  set(sh, 'A2', '100'); assert.equal(val(sh, 'B1'), 105); assert.equal(val(sh, 'A4'), 210);
  // whole column
  set(sh, 'C1', '=SUM(A:A)'); assert.equal(val(sh, 'C1'), 105 + 105 + 210);
  set(sh, 'A5', '1'); assert.equal(val(sh, 'C1'), 105 + 105 + 210 + 1);
});

test('cross-sheet references', () => {
  const wb = book(); const [s1, s2] = wb.sheets;
  set(s2, 'A1', '42'); set(s1, 'A1', '=Sheet2!A1*2');
  assert.equal(val(s1, 'A1'), 84);
  set(s2, 'A1', '1'); assert.equal(val(s1, 'A1'), 2);
  wb.transact('rename', () => wb.renameSheet(s2, 'Data Sheet'));
  assert.equal(s1.getCell(0, 0).f, "='Data Sheet'!A1*2");
  assert.equal(val(s1, 'A1'), 2);
  set(s1, 'B1', '=Nowhere!A1'); assert.equal(val(s1, 'B1'), '#REF!');
});

test('circular references', () => {
  const wb = book(); const sh = wb.sheets[0];
  set(sh, 'A1', '=B1'); set(sh, 'B1', '=A1');
  assert.equal(val(sh, 'A1'), '#CIRC!');
  set(sh, 'B1', '5'); assert.equal(val(sh, 'A1'), 5);
  set(sh, 'C1', '=C1+1'); assert.equal(val(sh, 'C1'), '#CIRC!');
});

test('functions: math and stats', () => {
  assert.equal(f('A1', '=SUM(1,2,3)'), 6);
  assert.equal(f('A1', '=ROUND(2.675,2)'), 2.68);
  assert.equal(f('A1', '=ROUND(-2.5,0)'), -3);
  assert.equal(f('A1', '=ROUND(1234.5678,-2)'), 1200);
  assert.equal(f('A1', '=ROUNDUP(3.14159,2)'), 3.15);
  assert.equal(f('A1', '=ROUNDDOWN(3.14159,2)'), 3.14);
  assert.equal(f('A1', '=MOD(-7,3)'), 2);
  assert.equal(f('A1', '=INT(-3.5)'), -4);
  assert.equal(f('A1', '=TRUNC(-3.5)'), -3);
  assert.equal(f('A1', '=CEILING(4.3,0.5)'), 4.5);
  assert.equal(f('A1', '=FLOOR(4.3,0.5)'), 4);
  assert.equal(f('A1', '=AVERAGE(1,2,3,4)'), 2.5);
  assert.equal(f('A1', '=MEDIAN(1,2,3,4)'), 2.5);
  assert.equal(f('A1', '=MAX(1,5,3)'), 5);
  assert.equal(f('A1', '=MIN(1,5,3)'), 1);
  assert.equal(f('A1', '=COUNT(1,"a",TRUE,"3")'), 3);
  assert.equal(f('A1', '=COUNTA(1,"a",TRUE,"")'), 4);
  approx(f('A1', '=STDEV(2,4,4,4,5,5,7,9)'), 2.138089935);
  approx(f('A1', '=STDEV.P(2,4,4,4,5,5,7,9)'), 2);
  assert.equal(f('A1', '=LARGE({3,1,2},2)'), 2);
  assert.equal(f('A1', '=SMALL({3,1,2},2)'), 2);
  assert.equal(f('A1', '=SUMPRODUCT({1,2,3},{4,5,6})'), 32);
  assert.equal(f('A1', '=PRODUCT(2,3,4)'), 24);
  assert.equal(f('A1', '=POWER(2,10)'), 1024);
  assert.equal(f('A1', '=SQRT(-1)'), '#NUM!');
  assert.equal(f('A1', '=ABS(-3)'), 3);
  assert.equal(f('A1', '=GCD(12,18)'), 6);
  assert.equal(f('A1', '=LCM(4,6)'), 12);
  assert.equal(f('A1', '=FACT(5)'), 120);
  assert.equal(f('A1', '=COMBIN(5,2)'), 10);
  assert.equal(f('A1', '=RANK(3,{1,2,3,4})'), 2);
  assert.equal(f('A1', '=PERCENTILE({1,2,3,4},0.5)'), 2.5);
  assert.equal(f('A1', '=QUARTILE({1,2,3,4,5},1)'), 2);
  assert.equal(f('A1', '=MODE(1,2,2,3)'), 2);
  assert.equal(f('A1', '=SUBTOTAL(9,{1,2,3})'), 6);
  approx(f('A1', '=NORM.S.DIST(1.96,TRUE)'), 0.975, 1e-4);
  approx(f('A1', '=NORM.S.INV(0.975)'), 1.959964, 1e-5);
  assert.equal(f('A1', '=ROMAN(1987)'), 'MCMLXXXVII');
  assert.equal(f('A1', '=SUM(A2:A3,5)'), 5);
});

test('functions: conditional aggregates', () => {
  const wb = book(); const sh = wb.sheets[0];
  const data = [['apple', 10], ['banana', 20], ['apple', 30], ['cherry', 40], ['Apple pie', 5]];
  wb.transact('fill', () => data.forEach((row, i) => { sh.setInput(i, 0, row[0]); sh.setInput(i, 1, String(row[1])); }));
  set(sh, 'D1', '=SUMIF(A1:A5,"apple",B1:B5)'); assert.equal(val(sh, 'D1'), 40);
  set(sh, 'D2', '=COUNTIF(A1:A5,"apple*")'); assert.equal(val(sh, 'D2'), 3);
  set(sh, 'D3', '=COUNTIF(B1:B5,">15")'); assert.equal(val(sh, 'D3'), 3);
  set(sh, 'D4', '=SUMIFS(B1:B5,A1:A5,"apple",B1:B5,">15")'); assert.equal(val(sh, 'D4'), 30);
  set(sh, 'D5', '=AVERAGEIF(B1:B5,"<>20")'); assert.equal(val(sh, 'D5'), 85 / 4);
  set(sh, 'D6', '=COUNTIFS(A1:A5,"<>apple",B1:B5,">=20")'); assert.equal(val(sh, 'D6'), 2);
  set(sh, 'D7', '=MAXIFS(B1:B5,A1:A5,"apple")'); assert.equal(val(sh, 'D7'), 30);
  set(sh, 'D8', '=COUNTBLANK(A1:A10)'); assert.equal(val(sh, 'D8'), 5);
  set(sh, 'D9', '=SUMIF(B1:B5,">=20")'); assert.equal(val(sh, 'D9'), 90);
});

test('functions: text', () => {
  assert.equal(f('A1', '=LEFT("hello",2)'), 'he');
  assert.equal(f('A1', '=RIGHT("hello",3)'), 'llo');
  assert.equal(f('A1', '=MID("hello",2,3)'), 'ell');
  assert.equal(f('A1', '=LEN("hello")'), 5);
  assert.equal(f('A1', '=UPPER("abc")'), 'ABC');
  assert.equal(f('A1', '=PROPER("hello world")'), 'Hello World');
  assert.equal(f('A1', '=TRIM("  a   b  ")'), 'a b');
  assert.equal(f('A1', '=CONCATENATE("a",1,TRUE)'), 'a1TRUE');
  assert.equal(f('A1', '=CONCAT("a","b")'), 'ab');
  assert.equal(f('A1', '=TEXTJOIN(", ",TRUE,"a","","b")'), 'a, b');
  assert.equal(f('A1', '=SUBSTITUTE("aaa","a","b",2)'), 'aba');
  assert.equal(f('A1', '=REPLACE("abcdef",2,3,"X")'), 'aXef');
  assert.equal(f('A1', '=FIND("c","abc")'), 3);
  assert.equal(f('A1', '=FIND("z","abc")'), '#VALUE!');
  assert.equal(f('A1', '=SEARCH("B","abc")'), 2);
  assert.equal(f('A1', '=SEARCH("b?","abcd")'), 2);
  assert.equal(f('A1', '=EXACT("a","A")'), false);
  assert.equal(f('A1', '=TEXT(1234.5,"#,##0.00")'), '1,234.50');
  assert.equal(f('A1', '=TEXT(0.25,"0%")'), '25%');
  assert.equal(f('A1', '=TEXT(DATE(2024,1,5),"yyyy-mm-dd")'), '2024-01-05');
  assert.equal(f('A1', '=VALUE("1,234")'), 1234);
  assert.equal(f('A1', '=VALUE("12%")'), 0.12);
  assert.equal(f('A1', '=REPT("ab",3)'), 'ababab');
  assert.equal(f('A1', '=CHAR(65)'), 'A');
  assert.equal(f('A1', '=CODE("A")'), 65);
  assert.equal(f('A1', '=DOLLAR(1234.567)'), '$1,234.57');
  assert.equal(f('A1', '=FIXED(1234.567,1)'), '1,234.6');
  assert.equal(f('A1', '=TEXTBEFORE("a-b-c","-")'), 'a');
  assert.equal(f('A1', '=TEXTAFTER("a-b-c","-",-1)'), 'c');
  assert.equal(f('A1', '=T(5)'), '');
  assert.equal(f('A1', '=N("x")'), 0);
  assert.equal(f('A1', '=LOWER(123)'), '123');
});

test('functions: logical and information', () => {
  assert.equal(f('A1', '=IF(1>2,"yes","no")'), 'no');
  assert.equal(f('A1', '=IF(TRUE,1)'), 1);
  assert.equal(f('A1', '=IF(FALSE,1)'), false);
  assert.equal(f('A1', '=IF(1>2,1/0,"safe")'), 'safe');
  assert.equal(f('A1', '=IFS(1>2,"a",2>1,"b")'), 'b');
  assert.equal(f('A1', '=IFS(1>2,"a")'), '#N/A');
  assert.equal(f('A1', '=IFERROR(1/0,"err")'), 'err');
  assert.equal(f('A1', '=IFERROR(5,"err")'), 5);
  assert.equal(f('A1', '=IFNA(NA(),"na")'), 'na');
  assert.equal(f('A1', '=AND(TRUE,1,"TRUE")'), true);
  assert.equal(f('A1', '=OR(FALSE,0)'), false);
  assert.equal(f('A1', '=XOR(TRUE,TRUE)'), false);
  assert.equal(f('A1', '=NOT(0)'), true);
  assert.equal(f('A1', '=SWITCH(2,1,"one",2,"two","other")'), 'two');
  assert.equal(f('A1', '=SWITCH(9,1,"one","other")'), 'other');
  assert.equal(f('A1', '=LET(x,2,y,3,x*y)'), 6);
  assert.equal(f('A1', '=ISNUMBER(1)'), true);
  assert.equal(f('A1', '=ISTEXT("a")'), true);
  assert.equal(f('A1', '=ISBLANK(B1)'), true);
  assert.equal(f('A1', '=ISERROR(1/0)'), true);
  assert.equal(f('A1', '=ISNA(NA())'), true);
  assert.equal(f('A1', '=ISERR(NA())'), false);
  assert.equal(f('A1', '=ISEVEN(4)'), true);
  assert.equal(f('A1', '=ISODD(4)'), false);
  assert.equal(f('A1', '=ERROR.TYPE(1/0)'), 2);
  assert.equal(f('A1', '=TYPE("a")'), 2);
  assert.equal(f('A1', '=FOO(1)'), '#NAME?');
  assert.equal(f('A1', '=SUM('), '#ERROR!');
});

test('functions: lookup and reference', () => {
  const wb = book(); const sh = wb.sheets[0];
  const data = [['id', 'name', 'score'], [1, 'ann', 90], [2, 'bob', 80], [3, 'cat', 70]];
  wb.transact('fill', () => data.forEach((row, i) => row.forEach((v, j) => sh.setInput(i, j, String(v)))));
  set(sh, 'E1', '=VLOOKUP(2,A2:C4,2,FALSE)'); assert.equal(val(sh, 'E1'), 'bob');
  set(sh, 'E2', '=VLOOKUP(9,A2:C4,2,FALSE)'); assert.equal(val(sh, 'E2'), '#N/A');
  set(sh, 'E3', '=VLOOKUP(2.5,A2:C4,3)'); assert.equal(val(sh, 'E3'), 80);
  set(sh, 'E4', '=HLOOKUP("name",A1:C4,3,FALSE)'); assert.equal(val(sh, 'E4'), 'bob');
  set(sh, 'E5', '=INDEX(A1:C4,2,2)'); assert.equal(val(sh, 'E5'), 'ann');
  set(sh, 'E6', '=MATCH("cat",B1:B4,0)'); assert.equal(val(sh, 'E6'), 4);
  set(sh, 'E7', '=INDEX(C1:C4,MATCH("bob",B1:B4,0))'); assert.equal(val(sh, 'E7'), 80);
  set(sh, 'E8', '=XLOOKUP("cat",B2:B4,C2:C4)'); assert.equal(val(sh, 'E8'), 70);
  set(sh, 'E9', '=XLOOKUP("zzz",B2:B4,C2:C4,"none")'); assert.equal(val(sh, 'E9'), 'none');
  set(sh, 'E10', '=OFFSET(A1,1,1)'); assert.equal(val(sh, 'E10'), 'ann');
  set(sh, 'E11', '=SUM(OFFSET(C1,1,0,3,1))'); assert.equal(val(sh, 'E11'), 240);
  set(sh, 'E12', '=INDIRECT("C"&3)'); assert.equal(val(sh, 'E12'), 80);
  set(sh, 'E13', '=ROW()'); assert.equal(val(sh, 'E13'), 13);
  set(sh, 'E14', '=COLUMN(C1)'); assert.equal(val(sh, 'E14'), 3);
  set(sh, 'E15', '=ROWS(A1:C4)'); assert.equal(val(sh, 'E15'), 4);
  set(sh, 'E16', '=ADDRESS(2,3)'); assert.equal(val(sh, 'E16'), '$C$2');
  set(sh, 'E17', '=CHOOSE(2,"a","b","c")'); assert.equal(val(sh, 'E17'), 'b');
  set(sh, 'E18', '=LOOKUP(2,A2:A4,B2:B4)'); assert.equal(val(sh, 'E18'), 'bob');
  set(sh, 'E19', '=SUM(A2:A4 )'); assert.equal(val(sh, 'E19'), 6);
  set(sh, 'E20', '=SUM(INDEX(A1:C4,0,3))'); assert.equal(val(sh, 'E20'), 240);
  set(sh, 'E21', '=XMATCH(80,C2:C4)'); assert.equal(val(sh, 'E21'), 2);
  // INDIRECT dependency updates
  set(sh, 'C3', '85'); assert.equal(val(sh, 'E12'), 85);
});

test('dynamic arrays spill', () => {
  const wb = book(); const sh = wb.sheets[0];
  set(sh, 'A1', '=SEQUENCE(3,2)');
  assert.equal(val(sh, 'A1'), 1); assert.equal(val(sh, 'B1'), 2); assert.equal(val(sh, 'A3'), 5); assert.equal(val(sh, 'B3'), 6);
  assert.deepEqual(sh.getCell(0, 0).spill, { r1: 0, c1: 0, r2: 2, c2: 1 });
  set(sh, 'C1', '=SUM(A1:B3)'); assert.equal(val(sh, 'C1'), 21);
  set(sh, 'D1', '=A3*10'); assert.equal(val(sh, 'D1'), 50);
  // blocked spill
  set(sh, 'B2', 'x'); assert.equal(val(sh, 'A1'), '#SPILL!'); assert.equal(val(sh, 'A2'), null);
  set(sh, 'B2', ''); assert.equal(val(sh, 'A1'), 1); assert.equal(val(sh, 'B2'), 4); assert.equal(val(sh, 'D1'), 50);
  // shrink spill clears cells
  set(sh, 'A1', '=SEQUENCE(2,1)'); assert.equal(val(sh, 'B1'), null); assert.equal(val(sh, 'A3'), null); assert.equal(val(sh, 'C1'), 3);
  // array operators
  set(sh, 'F1', '={1,2,3}*2'); assert.equal(val(sh, 'H1'), 6);
  set(sh, 'F2', '=SUM({1,2,3}*{1,2,3})'); assert.equal(val(sh, 'F2'), 14);
  set(sh, 'F3', '=TRANSPOSE({1,2,3})'); assert.equal(val(sh, 'F5'), 3);
  set(sh, 'G3', '=FILTER({1;2;3;4},{1;2;3;4}>2)'); assert.equal(val(sh, 'G3'), 3); assert.equal(val(sh, 'G4'), 4);
  set(sh, 'H3', '=UNIQUE({1;2;2;3})'); assert.equal(val(sh, 'H5'), 3);
  set(sh, 'I3', '=SORT({3;1;2},1,-1)'); assert.equal(val(sh, 'I3'), 3); assert.equal(val(sh, 'I5'), 1);
  set(sh, 'J1', '=ABS({-1,-2})'); assert.equal(val(sh, 'K1'), 2);
  set(sh, 'J2', '=IF({TRUE,FALSE},"y","n")'); assert.equal(val(sh, 'K2'), 'n');
  set(sh, 'L1', '=SUM(IF(A1:A2>1,1,0))'); assert.equal(val(sh, 'L1'), 1);
});

test('functions: dates and times', () => {
  const d = dateToSerial(2024, 1, 15);
  assert.equal(f('A1', '=DATE(2024,1,15)'), d);
  assert.equal(f('A1', '=YEAR(DATE(2024,1,15))'), 2024);
  assert.equal(f('A1', '=MONTH(DATE(2024,1,15))'), 1);
  assert.equal(f('A1', '=DAY(DATE(2024,1,15))'), 15);
  assert.equal(f('A1', '=WEEKDAY(DATE(2024,1,15))'), 2); // Monday
  assert.equal(f('A1', '=WEEKDAY(DATE(2024,1,15),2)'), 1);
  assert.equal(f('A1', '=EOMONTH(DATE(2024,1,15),1)'), dateToSerial(2024, 2, 29));
  assert.equal(f('A1', '=EDATE(DATE(2024,1,31),1)'), dateToSerial(2024, 2, 29));
  assert.equal(f('A1', '=DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"Y")'), 4);
  assert.equal(f('A1', '=DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"M")'), 49);
  assert.equal(f('A1', '=DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"D")'), 1516);
  assert.equal(f('A1', '=DAYS(DATE(2024,3,1),DATE(2024,2,1))'), 29);
  assert.equal(f('A1', '=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))'), 23);
  assert.equal(f('A1', '=WORKDAY(DATE(2024,1,5),1)'), dateToSerial(2024, 1, 8));
  assert.equal(f('A1', '=HOUR(TIME(13,45,10))'), 13);
  assert.equal(f('A1', '=MINUTE(TIME(13,45,10))'), 45);
  assert.equal(f('A1', '=SECOND(TIME(13,45,10))'), 10);
  approx(f('A1', '=TIME(12,0,0)'), 0.5);
  assert.equal(f('A1', '=DATEVALUE("2024-01-15")'), d);
  assert.equal(f('A1', '=DATE(2024,13,1)'), dateToSerial(2025, 1, 1));
  assert.equal(f('A1', '=WEEKNUM(DATE(2024,1,15))'), 3);
  assert.equal(f('A1', '=ISOWEEKNUM(DATE(2024,1,1))'), 1);
  assert.equal(f('A1', '=YEAR("2024-05-05")'), 2024);
  assert.equal(f('A1', '=DAYS360(DATE(2024,1,1),DATE(2024,12,31))'), 360);
  approx(f('A1', '=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1))'), 0.5);
  const wb = book(); const sh = wb.sheets[0];
  set(sh, 'A1', '=TODAY()');
  assert.equal(sh.getCell(0, 0).autoFmt, 'm/d/yyyy');
  set(sh, 'A2', '=A1+1'); assert.equal(sh.getCell(1, 0).autoFmt, 'm/d/yyyy');
  set(sh, 'A3', '2024-01-05'); assert.equal(val(sh, 'A3'), dateToSerial(2024, 1, 5)); assert.equal(sh.getStyle(2, 0).numFmt, 'yyyy-mm-dd');
});

test('functions: financial', () => {
  approx(f('A1', '=PMT(0.05/12,360,200000)'), -1073.643246, 1e-5);
  approx(f('A1', '=PV(0.08/12,120,-1000)'), 82421.48, 1e-2);
  approx(f('A1', '=FV(0.06/12,120,-100)'), 16387.93, 1e-2);
  approx(f('A1', '=NPER(0.01,-100,1000)'), 10.588, 1e-3);
  approx(f('A1', '=RATE(360,-1073.643246,200000)*12'), 0.05, 1e-6);
  approx(f('A1', '=NPV(0.1,-1000,500,500,500)'), 221.29636, 1e-3);
  approx(f('A1', '=IRR({-1000,500,500,500})'), 0.23375, 1e-4);
  approx(f('A1', '=IPMT(0.1/12,1,36,8000)'), -66.6667, 1e-3);
  approx(f('A1', '=PPMT(0.1/12,1,36,8000)'), -191.4703, 1e-3);
  approx(f('A1', '=SLN(30000,7500,10)'), 2250);
  approx(f('A1', '=SYD(30000,7500,10,1)'), 4090.909, 1e-3);
  approx(f('A1', '=DDB(2400,300,10,1)'), 480);
  approx(f('A1', '=EFFECT(0.0525,4)'), 0.053543, 1e-6);
});

test('functions: engineering and misc', () => {
  assert.equal(f('A1', '=DEC2BIN(10)'), '1010');
  assert.equal(f('A1', '=BIN2DEC("1010")'), 10);
  assert.equal(f('A1', '=DEC2HEX(255)'), 'FF');
  assert.equal(f('A1', '=HEX2DEC("FF")'), 255);
  assert.equal(f('A1', '=BITAND(12,10)'), 8);
  assert.equal(f('A1', '=BITXOR(12,10)'), 6);
  approx(f('A1', '=CONVERT(1,"mi","km")'), 1.609344);
  approx(f('A1', '=CONVERT(100,"C","F")'), 212);
  assert.equal(f('A1', '=CONVERT(1,"mi","kg")'), '#N/A');
  assert.equal(f('A1', '=BASE(255,16)'), 'FF');
});

test('names', () => {
  const wb = book(); const sh = wb.sheets[0];
  set(sh, 'A1', '5'); set(sh, 'A2', '7');
  wb.transact('name', () => wb.setName('Rate', '=Sheet1!$A$1'));
  set(sh, 'B1', '=Rate*2'); assert.equal(val(sh, 'B1'), 10);
  set(sh, 'A1', '6'); assert.equal(val(sh, 'B1'), 12);
  wb.transact('name', () => wb.setName('Data', '=Sheet1!$A$1:$A$2'));
  set(sh, 'B2', '=SUM(Data)'); assert.equal(val(sh, 'B2'), 13);
  wb.transact('name', () => wb.setName('TaxRate', '=0.2'));
  set(sh, 'B3', '=100*TaxRate'); assert.equal(val(sh, 'B3'), 20);
  set(sh, 'B4', '=Unknown+1'); assert.equal(val(sh, 'B4'), '#NAME?');
});

test('reference shifting (copy semantics)', () => {
  assert.equal(shiftFormula('=A1+$B$2+C$3+$D4', 2, 3), '=D3+$B$2+F$3+$D6');
  assert.equal(shiftFormula('=SUM(A1:B2)', 1, 0), '=SUM(A2:B3)');
  assert.equal(shiftFormula('=A1', -1, 0), '=#REF!');
  assert.equal(shiftFormula('=Sheet2!A1', 1, 1), '=Sheet2!B2');
  assert.equal(shiftFormula('=A:A', 0, 1), '=B:B');
  assert.equal(shiftFormula('=1:1', 1, 0), '=2:2');
  assert.equal(shiftFormula('="A1"&A1', 0, 1), '="A1"&B1');
});

test('reference adjustment for insert/delete', () => {
  assert.equal(adjustFormulaForInsertDelete('=A1+A5+SUM(A2:A4)', 'S', 'S', 'row', 2, 1), '=A1+A6+SUM(A2:A5)');
  assert.equal(adjustFormulaForInsertDelete('=A1+A5+SUM(A2:A4)', 'S', 'S', 'row', 2, -1), '=A1+A4+SUM(A2:A3)');
  assert.equal(adjustFormulaForInsertDelete('=A3', 'S', 'S', 'row', 2, -1), '=#REF!');
  assert.equal(adjustFormulaForInsertDelete('=SUM(A2:A4)', 'S', 'S', 'row', 1, -3), '=SUM(#REF!)');
  assert.equal(adjustFormulaForInsertDelete('=SUM(B1:D1)', 'S', 'S', 'col', 2, 2), '=SUM(B1:F1)');
  assert.equal(adjustFormulaForInsertDelete('=Other!A5', 'S', 'S', 'row', 0, 1), '=Other!A5');
  assert.equal(adjustFormulaForInsertDelete('=Other!A5', 'S', 'Other', 'row', 0, 1), '=Other!A6');
  assert.equal(adjustFormulaForInsertDelete('=SUM(A:A)', 'S', 'S', 'row', 0, 5), '=SUM(A:A)');
});

test('model: insert/delete rows and columns update formulas and cells', () => {
  const wb = book(); const sh = wb.sheets[0];
  set(sh, 'A1', '1'); set(sh, 'A2', '2'); set(sh, 'A3', '3'); set(sh, 'B1', '=SUM(A1:A3)'); set(sh, 'B2', '=A3');
  wb.transact('ins', () => sh.insertRows(1, 2));
  assert.equal(sh.getCell(3, 0).v, 2); assert.equal(sh.getCell(4, 0).v, 3);
  assert.equal(sh.getCell(0, 1).f, '=SUM(A1:A5)'); assert.equal(sh.getCell(3, 1).f, '=A5');
  assert.equal(val(sh, 'B1'), 6); assert.equal(val(sh, 'B4'), 3);
  wb.transact('del', () => sh.deleteRows(1, 2));
  assert.equal(sh.getCell(1, 0).v, 2); assert.equal(sh.getCell(0, 1).f, '=SUM(A1:A3)'); assert.equal(val(sh, 'B1'), 6);
  wb.transact('delcol', () => sh.deleteCols(0, 1));
  assert.equal(sh.getCell(0, 0).f, '=SUM(#REF!)'); assert.equal(val(sh, 'A1'), '#REF!');
  wb.undo(); assert.equal(sh.getCell(0, 1).f, '=SUM(A1:A3)'); assert.equal(val(sh, 'B1'), 6); assert.equal(sh.getCell(0, 0).v, 1);
  wb.transact('inscol', () => sh.insertCols(0, 1));
  assert.equal(sh.getCell(0, 2).f, '=SUM(B1:B3)'); assert.equal(val(sh, 'C1'), 6);
});

test('model: undo/redo', () => {
  const wb = book(); const sh = wb.sheets[0];
  set(sh, 'A1', '1'); set(sh, 'A1', '2'); set(sh, 'B1', '=A1*2');
  assert.equal(val(sh, 'B1'), 4);
  wb.undo(); assert.equal(sh.getCell(0, 1), null);
  wb.undo(); assert.equal(val(sh, 'A1'), 1);
  wb.redo(); assert.equal(val(sh, 'A1'), 2);
  wb.redo(); assert.equal(val(sh, 'B1'), 4);
  assert.equal(wb.canRedo(), false);
  wb.transact('fmt', () => sh.applyStyle({ r1: 0, c1: 0, r2: 0, c2: 0 }, { bold: true }));
  assert.equal(sh.getStyle(0, 0).bold, true);
  wb.undo(); assert.equal(sh.getStyle(0, 0).bold, undefined);
  wb.transact('width', () => sh.setColWidth(0, 150)); assert.equal(sh.colWidth(0), 150);
  wb.undo(); assert.equal(sh.colWidth(0), 80);
  wb.transact('merge', () => sh.merge({ r1: 0, c1: 0, r2: 1, c2: 1 })); assert.equal(sh.merges.length, 1);
  wb.undo(); assert.equal(sh.merges.length, 0);
});

test('model: sheets add/remove/rename with undo', () => {
  const wb = book();
  wb.transact('add', () => wb.addSheet('Extra'));
  assert.equal(wb.sheets.length, 3);
  set(wb.sheets[0], 'A1', '=Extra!A1+1'); set(wb.sheets[2], 'A1', '9'); assert.equal(val(wb.sheets[0], 'A1'), 10);
  wb.transact('rm', () => wb.removeSheet(wb.sheets[2]));
  assert.equal(wb.sheets.length, 2); assert.equal(wb.sheets[0].getCell(0, 0).f, '=#REF!+1');
  wb.undo(); assert.equal(wb.sheets.length, 3); assert.equal(val(wb.sheets[0], 'A1'), 10);
});

test('model: sort range', () => {
  const wb = book(); const sh = wb.sheets[0];
  const rows = [['name', 'n'], ['c', 3], ['a', 1], ['b', 2]];
  wb.transact('fill', () => rows.forEach((r, i) => r.forEach((v, j) => sh.setInput(i, j, String(v)))));
  set(sh, 'C2', '=B2*10'); set(sh, 'C3', '=B3*10'); set(sh, 'C4', '=B4*10');
  wb.transact('sort', () => sh.sortRange({ r1: 0, c1: 0, r2: 3, c2: 2 }, [{ col: 0, asc: true }], true));
  assert.equal(val(sh, 'A2'), 'a'); assert.equal(val(sh, 'B2'), 1); assert.equal(val(sh, 'C2'), 10); assert.equal(sh.getCell(1, 2).f, '=B2*10');
  assert.equal(val(sh, 'A4'), 'c'); assert.equal(val(sh, 'C4'), 30);
  wb.undo(); assert.equal(val(sh, 'A2'), 'c');
});

test('model: filter hides rows', () => {
  const wb = book(); const sh = wb.sheets[0];
  const rows = [['fruit', 'n'], ['apple', 3], ['pear', 1], ['apple', 2]];
  wb.transact('fill', () => rows.forEach((r, i) => r.forEach((v, j) => sh.setInput(i, j, String(v)))));
  wb.transact('filter', () => sh.setFilter({ r1: 0, c1: 0, r2: 3, c2: 1, conds: { 0: { values: ['apple'] } } }));
  assert.equal(sh.rowAxis.filtered.has(2), true); assert.equal(sh.rowAxis.filtered.has(1), false);
  set(sh, 'D1', '=SUBTOTAL(109,B2:B4)'); assert.equal(val(sh, 'D1'), 5);
  assert.equal(sh.rowAxis.offset(3), 40);
});

test('model: JSON round trip', () => {
  const wb = book(); const sh = wb.sheets[0];
  set(sh, 'A1', '1'); set(sh, 'A2', '=A1+1'); set(sh, 'B1', 'text');
  wb.transact('fmt', () => { sh.applyStyle({ r1: 0, c1: 0, r2: 0, c2: 0 }, { bold: true, numFmt: '0.00' }); sh.setColWidth(1, 120); sh.merge({ r1: 2, c1: 0, r2: 3, c2: 1 }); });
  const j = JSON.parse(JSON.stringify(wb.toJSON()));
  const wb2 = Workbook.fromJSON(j);
  const s2 = wb2.sheets[0];
  assert.equal(val(s2, 'A2'), 2); assert.equal(val(s2, 'B1'), 'text'); assert.equal(s2.getStyle(0, 0).bold, true); assert.equal(s2.colWidth(1), 120); assert.equal(s2.merges.length, 1);
});

test('axis offsets with hidden and custom sizes', () => {
  const wb = book(); const sh = wb.sheets[0];
  wb.transact('x', () => { sh.setRowHeight(2, 40); sh.setRowHidden(5, true); });
  assert.equal(sh.rowAxis.offset(0), 0); assert.equal(sh.rowAxis.offset(3), 80); assert.equal(sh.rowAxis.offset(6), 120); assert.equal(sh.rowAxis.offset(7), 140);
  assert.equal(sh.rowAxis.indexAt(85, 1000), 3); assert.equal(sh.rowAxis.indexAt(120, 1000), 6); assert.equal(sh.rowAxis.indexAt(45, 1000), 2);
});
