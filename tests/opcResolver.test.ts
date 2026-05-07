import { describe, expect, it } from 'vitest';

import {
  relsPathFor,
  resolvePackagePaths,
  resolveWorkbookPath,
} from '../src/opcResolver.js';

const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const rootRels = (target: string): string => `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="${target}"/>
</Relationships>`;

const workbookXml = (rIds: string[]): string => `<workbook xmlns:r="${REL_NS}">
  <sheets>
    ${rIds.map((id, i) => `<sheet name="S${i + 1}" sheetId="${i + 1}" r:id="${id}"/>`).join('\n    ')}
  </sheets>
</workbook>`;

const workbookRels = (rels: { id: string; type: string; target: string }[]): string =>
  `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${rels.map((r) => `<Relationship Id="${r.id}" Type="${REL_NS}/${r.type}" Target="${r.target}"/>`).join('\n  ')}
</Relationships>`;

describe('relsPathFor', () => {
  it('returns _rels/<name>.rels at the package root', () => {
    expect(relsPathFor('foo.xml')).toBe('_rels/foo.xml.rels');
  });

  it('places the rels file under the part directory', () => {
    expect(relsPathFor('xl/workbook.xml')).toBe('xl/_rels/workbook.xml.rels');
    expect(relsPathFor('xl/worksheets/sheet1.xml')).toBe(
      'xl/worksheets/_rels/sheet1.xml.rels',
    );
  });
});

describe('resolveWorkbookPath', () => {
  it('resolves an absolute Target by stripping the leading slash', () => {
    expect(resolveWorkbookPath(rootRels('/xl/workbook.xml'))).toBe('xl/workbook.xml');
  });

  it('resolves a relative Target as root-relative', () => {
    expect(resolveWorkbookPath(rootRels('xl/workbook.xml'))).toBe('xl/workbook.xml');
  });

  it('returns undefined when no officeDocument relationship is present', () => {
    const xml = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="${REL_NS}/extended-properties" Target="docProps/app.xml"/>
    </Relationships>`;
    expect(resolveWorkbookPath(xml)).toBeUndefined();
  });

  it('ignores external-mode relationships', () => {
    const xml = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="http://elsewhere/wb.xml" TargetMode="External"/>
    </Relationships>`;
    expect(resolveWorkbookPath(xml)).toBeUndefined();
  });
});

describe('resolvePackagePaths', () => {
  it('resolves a workbook in the conventional xl/ layout', () => {
    const paths = resolvePackagePaths(
      rootRels('/xl/workbook.xml'),
      workbookXml(['rId1']),
      workbookRels([
        { id: 'rId1', type: 'worksheet', target: 'worksheets/sheet1.xml' },
        { id: 'rId2', type: 'sharedStrings', target: 'sharedStrings.xml' },
        { id: 'rId3', type: 'styles', target: 'styles.xml' },
      ]),
    );
    expect(paths).toBeDefined();
    expect(paths?.workbook).toBe('xl/workbook.xml');
    expect(paths?.sharedStrings).toBe('xl/sharedStrings.xml');
    expect(paths?.styles).toBe('xl/styles.xml');
    expect(paths?.sheetByRId.get('rId1')).toBe('xl/worksheets/sheet1.xml');
  });

  it('resolves a non-Excel layout where parts live outside xl/', () => {
    const paths = resolvePackagePaths(
      rootRels('/spreadsheet/wb.xml'),
      workbookXml(['rIdA']),
      workbookRels([
        { id: 'rIdA', type: 'worksheet', target: 'sheets/main.xml' },
        { id: 'rIdB', type: 'sharedStrings', target: '../strings/sst.xml' },
      ]),
    );
    expect(paths?.workbook).toBe('spreadsheet/wb.xml');
    expect(paths?.sharedStrings).toBe('strings/sst.xml');
    expect(paths?.sheetByRId.get('rIdA')).toBe('spreadsheet/sheets/main.xml');
  });

  it('omits sharedStrings/styles when the workbook does not declare them', () => {
    const paths = resolvePackagePaths(
      rootRels('/xl/workbook.xml'),
      workbookXml(['rId1']),
      workbookRels([
        { id: 'rId1', type: 'worksheet', target: 'worksheets/sheet1.xml' },
      ]),
    );
    expect(paths?.sharedStrings).toBeUndefined();
    expect(paths?.styles).toBeUndefined();
  });

  it('skips dangling rIds (workbook references a sheet that has no matching rel)', () => {
    const paths = resolvePackagePaths(
      rootRels('/xl/workbook.xml'),
      workbookXml(['rId1', 'rIdGhost']),
      workbookRels([
        { id: 'rId1', type: 'worksheet', target: 'worksheets/sheet1.xml' },
      ]),
    );
    expect(paths?.sheetByRId.size).toBe(1);
    expect(paths?.sheetByRId.has('rIdGhost')).toBe(false);
  });
});
