import { afterEach, describe, expect, it } from 'vitest';
import { analyzeFormula, relativePattern } from './formulaAnalysis';
import { analyzeBlock, brokenRefsReport, emptySheetScan, explainWorkbook, hardcodedReport, scanReport, type WorkbookScan } from './workbookScan';
import { diagnoseLookupNA, expandLet, planBreakLinks, planMoveLiteralsToInputs, replaceLiteral, toLet } from './formulaTools';
import { detectInspection, inspectionWrites } from './inspectionLanes';
import { listUserTemplates, saveUserTemplate } from './userTemplates';

// T_Formulas in CELLIX_TEST_DATA.xlsx
const T_FORMULAS_F = [
  ['Amount', 'Formula'],
  [10000, '=A2*18%'],
  [20000, '=A3*1'],
  [30000, '=ROUND(A4,2)'],
  [40000, '=A5*0.075'],
  [50000, '=#REF!+1'],
  [60000, '=IF(A7>50000,A7*0.1,0)'],
  [70000, '=SUM(A2:A7)'],
];
const T_FORMULAS_V = [
  ['Amount', 'Formula'],
  [10000, 1800],
  [20000, 20000],
  [30000, 30000],
  [40000, 3000],
  [50000, '#REF!'],
  [60000, 6000],
  [70000, 210000],
];

function scanOf(formulas: unknown[][], values: unknown[][], name = 'T_Formulas'): WorkbookScan {
  const sheet = emptySheetScan(name, 'Visible');
  sheet.rows = formulas.length;
  sheet.cols = 2;
  const graph = new Map<string, Set<string>>();
  const ext = new Set<string>();
  analyzeBlock(sheet, formulas, values, 0, 0, graph, ext);
  return {
    sheets: [sheet, { ...emptySheetScan('Hidden one', 'Hidden'), rows: 1, cols: 1 }],
    names: [{ name: 'BadName', formula: '=#REF!', broken: true }],
    graph: Object.fromEntries([...graph].map(([k, v]) => [k, [...v]])),
    externalWorkbooks: [...ext],
  };
}

describe('P3.5 hard-coded literal detection', () => {
  it.each([
    ['=A2*18%', ['18%']],
    ['=A3*1', []],
    ['=ROUND(A4,2)', []],
    ['=A5*0.075', ['0.075']],
    ['=IF(A7>50000,A7*0.1,0)', ['50000', '0.1']],
    ['=SUM(A2:A7)', []],
    ['=VLOOKUP(A2,$G$2:$H$4,2,FALSE)', []],
    ["='Other Sheet'!B2*12", []],
  ])('%s → %j', (f, lits) => expect(analyzeFormula(f).literals.map((l) => l.text)).toEqual(lits));

  it('finds #REF!, external workbooks and cross-sheet refs', () => {
    expect(analyzeFormula('=#REF!+1').hasRefError).toBe(true);
    expect(analyzeFormula("='[Budget.xlsx]Sheet1'!A1*2").externalWorkbooks).toEqual(['Budget.xlsx']);
    expect(analyzeFormula("=SUM('T_Days'!D2:D6)+Inputs!B2").sheetRefs.sort()).toEqual(['Inputs', 'T_Days']);
  });

  it('relative pattern is the same for a filled-down formula', () => {
    expect(relativePattern('=A2*2', 1, 1)).toBe(relativePattern('=A3*2', 2, 1));
    expect(relativePattern('=A2*2', 1, 1)).not.toBe(relativePattern('=A2*3', 2, 1));
  });
});

describe('P2.4 / P3.5 / P3.6 / P3.22 reports on T_Formulas', () => {
  const scan = scanOf(T_FORMULAS_F, T_FORMULAS_V);
  it('flags B2, B5, B7 and not B3, B4, B8', () => {
    const cells = scan.sheets[0].hardcoded.map((c) => c.address);
    expect(cells).toEqual(['B2', 'B5', 'B7']);
    expect(hardcodedReport(scan)).toMatch(/3 formula\(s\) contain hard-coded numbers/);
  });
  it('finds B6 #REF! and the broken name', () => {
    expect(scan.sheets[0].refErrors.map((c) => c.address)).toEqual(['B6']);
    const r = brokenRefsReport(scan);
    expect(r).toMatch(/T_Formulas!B6/);
    expect(r).toMatch(/BadName/);
  });
  it('scan report lists every kind and the hidden sheet, read-only', () => {
    const r = scanReport(scan);
    expect(r).toMatch(/Read-only: nothing was changed/);
    expect(r).toMatch(/Hard-coded numbers in formulas\*\* — 3/);
    expect(r).toMatch(/1 hidden sheet\(s\): Hidden one/);
  });
  it('workbook explanation covers hidden sheets and links', () => {
    const r = explainWorkbook(scan);
    expect(r).toMatch(/Hidden one.*hidden/);
    expect(r).toMatch(/Cross-sheet links/);
  });
});

describe('P3.5 fix: move constants to inputs', () => {
  it('replaces only the literal, never a cell reference', () => {
    expect(replaceLiteral('=A2*18%', '18%', 'X')).toBe('=A2*X');
    expect(replaceLiteral('=A18*18', '18', 'X')).toBe('=A18*X');
  });
  it('writes labelled inputs and repoints the formula', () => {
    const scan = scanOf(T_FORMULAS_F, T_FORMULAS_V);
    const plan = planMoveLiteralsToInputs(scan.sheets[0].hardcoded);
    expect(plan.actions[0]).toMatchObject({ type: 'ADD_SHEET', name: 'CELLIX Inputs' });
    const formulas = (plan.actions[2] as { operations: Array<{ address: string; formula: string }> }).operations;
    expect(formulas.find((o) => o.address === 'B2')?.formula).toBe("=A2*'CELLIX Inputs'!$B$2");
    expect(formulas.find((o) => o.address === 'B7')?.formula).toBe("=IF(A7>'CELLIX Inputs'!$B$4,A7*'CELLIX Inputs'!$B$5,0)");
  });
});

describe('P3.6 fix: break links', () => {
  it('replaces #REF! formulas with a marked text value', () => {
    const plan = planBreakLinks(scanOf(T_FORMULAS_F, T_FORMULAS_V));
    expect(plan.actions[0]).toMatchObject({ type: 'BATCH_SET', sheetName: 'T_Formulas', operations: [{ address: 'B6', value: '#REF! (was =#REF!+1)' }] });
  });
});

describe('P3.13 / P4.6 lookup #N/A (T_Lookup)', () => {
  const F = [
    ['Code(text)', 'Amount', 'Name via VLOOKUP', null, null, null, 'Master Code(number)', 'Master Name'],
    ['101', 100, '=VLOOKUP(A2,$G$2:$H$4,2,FALSE)', null, null, null, 101, 'ABC Traders'],
    ['102', 200, '=VLOOKUP(A3,$G$2:$H$4,2,FALSE)', null, null, null, 102, 'XYZ Pvt Ltd'],
    ['103', 300, '=VLOOKUP(A4,$G$2:$H$4,2,FALSE)', null, null, null, 103, 'Mehta & Sons'],
  ];
  const V = F.map((r, i) => (i === 0 ? r : [r[0], r[1], '#N/A', null, null, null, r[6], r[7]]));
  it('diagnoses text keys vs number master codes', () => {
    const d = diagnoseLookupNA(F, V, 2)!;
    expect(d.kind).toBe('text-vs-number');
    expect(d.message).toMatch(/column A are \*\*text\*\*.*column G are \*\*numbers\*\*/);
    expect(d.fix).toEqual([
      { address: 'A2', value: 101 },
      { address: 'A3', value: 102 },
      { address: 'A4', value: 103 },
    ]);
  });
});

describe('P3.17 LET rewrite', () => {
  it('names the repeated reference and the as-on date; expands back exactly', () => {
    const f = '=IF(ISNUMBER(C2),DATE(2026,3,31)-C2,"")';
    const r = toLet(f, { C2: 'invoiceDate' })!;
    expect(r.rewritten).toBe('=LET(asOn,DATE(2026,3,31),invoiceDate,C2,IF(ISNUMBER(invoiceDate),asOn-invoiceDate,""))');
    expect(r.equivalent).toBe(true);
    expect(expandLet(r.rewritten)).toBe(f);
  });
  it('returns null when nothing is worth naming', () => {
    expect(toLet('=A2+1')).toBeNull();
  });
});

describe('intent routing', () => {
  it.each([
    ['Scan this workbook for problems', 'scan'],
    ['Find hard-coded numbers in my formulas', 'hardcoded'],
    ['Move the hard-coded numbers to input cells', 'hardcoded-fix'],
    ['Check for broken references in this workbook', 'broken-refs'],
    ['Break the broken links', 'break-links'],
    ['Why is column C showing #N/A?', 'lookup-why'],
    ['Fix the #N/A in column C', 'lookup-fix'],
    ['Rewrite my formula in F2 using LET', 'let'],
    ['Explain this workbook', 'explain-workbook'],
    ['Save this sheet as a template', 'save-template'],
  ])('%s → %s', (m, k) => expect(detectInspection(m)).toBe(k));

  it.each(['Explain this formula', 'Put the number 1 in cells A1 to A5', 'Sort by Amount descending', 'Validate the GSTINs in column B'])(
    'ignores %s',
    (m) => expect(detectInspection(m)).toBeNull(),
  );

  it('only fix lanes write', () => {
    expect(inspectionWrites('scan')).toBe(false);
    expect(inspectionWrites('lookup-fix')).toBe(true);
  });
});

describe('P3.24 user templates store headers only', () => {
  const store: Record<string, string> = {};
  (globalThis as unknown as { localStorage: Storage }).localStorage = {
    getItem: (k: string) => store[k] ?? null,
    setItem: (k: string, v: string) => {
      store[k] = v;
    },
  } as Storage;
  afterEach(() => Object.keys(store).forEach((k) => delete store[k]));

  it('saves and lists; nothing but name + headers', () => {
    expect(saveUserTemplate({ name: 'My GST working', headers: ['GSTIN', 'Invoice', ''] })).toBe(true);
    expect(listUserTemplates()).toEqual([{ name: 'My GST working', headers: ['GSTIN', 'Invoice'] }]);
    expect(Object.keys(JSON.parse(store['cellix.userTemplates.v1'])[0])).toEqual(['name', 'headers']);
  });
});
