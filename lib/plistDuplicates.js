const PLIST_COLUMNS = ['code', 'pqFormDesc', 'pdesc1', 'type', 'width', 'height', 'length', 'thickness', 'materialWidth'];

function cell(row, index) {
  return index < row.length ? String(row[index] ?? '').trim() : '';
}

export function normalizePlistCode(code) {
  return String(code ?? '').trim().toUpperCase();
}

export function plistRowToRecord(row, rowIndex) {
  return {
    rowIndex,
    code: cell(row, 0),
    codeKey: normalizePlistCode(cell(row, 0)),
    pqFormDesc: cell(row, 1),
    pdesc1: cell(row, 2),
    type: cell(row, 3),
    width: cell(row, 4),
    height: cell(row, 5),
    length: cell(row, 6),
    thickness: cell(row, 7),
    materialWidth: cell(row, 8),
  };
}

export function recordSignature(rec, { ignoreMaterialWidth = false } = {}) {
  const keys = ignoreMaterialWidth
    ? ['pqFormDesc', 'pdesc1', 'type', 'width', 'height', 'length', 'thickness']
    : ['pqFormDesc', 'pdesc1', 'type', 'width', 'height', 'length', 'thickness', 'materialWidth'];
  return keys.map((k) => rec[k] || '').join('|');
}

export function scorePlistRecord(rec) {
  let score = 0;
  if (rec.materialWidth) score += 20;
  if (/^[\d.]+x[\d.]+x[\d.]+\s+[\u4e00-\u9fff]/.test(rec.pqFormDesc)) score += 15;
  if (['企筒', '地槽', '鐵角', '批灰角', 'W角', '闊槽', 'C槽', 'CT企筒打孔'].includes(rec.type)) score += 12;
  else if (rec.type && !/[()（）]/.test(rec.type)) score += 5;
  if (rec.width && rec.height && rec.length && rec.thickness) score += 10;
  if (rec.pqFormDesc) score += 2;
  if (rec.pdesc1) score += 1;
  return score;
}

export function diffRecords(a, b) {
  const diffs = [];
  for (const col of PLIST_COLUMNS.slice(1)) {
    const av = a[col] || '';
    const bv = b[col] || '';
    if (av !== bv) diffs.push({ column: col, a: av, b: bv });
  }
  return diffs;
}

export function analyzePlistDuplicates(plistRows) {
  const records = plistRows.map((row, i) => plistRowToRecord(row, i + 2));
  const byCode = new Map();

  for (const rec of records) {
    if (!rec.codeKey) continue;
    if (!byCode.has(rec.codeKey)) byCode.set(rec.codeKey, []);
    byCode.get(rec.codeKey).push(rec);
  }

  const duplicateGroups = [];
  let exactDuplicateRows = 0;
  let conflictingGroups = 0;
  let complementaryGroups = 0;

  for (const [codeKey, group] of byCode.entries()) {
    if (group.length < 2) continue;

    const signatures = new Map();
    for (const rec of group) {
      const sig = recordSignature(rec);
      if (!signatures.has(sig)) signatures.set(sig, []);
      signatures.get(sig).push(rec);
    }

    const sigEntries = [...signatures.entries()];
    const isExactDuplicate = sigEntries.length === 1;
    const dimSig = (rec) => [rec.width, rec.height, rec.length, rec.thickness, rec.materialWidth].join('|');
    const hasConflict = new Set(group.map(dimSig)).size > 1;

    const onlyMaterialWidthDiff = !hasConflict && sigEntries.length > 1 && sigEntries.every(([, rows]) => {
      const base = group[0];
      return diffRecords(base, rows[0]).every((d) => ['materialWidth', 'pqFormDesc', 'pdesc1'].includes(d.column));
    });
    const onlyDescDiff = !hasConflict && sigEntries.length > 1 && sigEntries.every(([, rows]) => {
      const base = group[0];
      return diffRecords(base, rows[0]).every((d) => ['pqFormDesc', 'pdesc1'].includes(d.column));
    });

    if (isExactDuplicate) exactDuplicateRows += group.length - 1;
    else if (hasConflict) conflictingGroups += 1;
    else if (onlyMaterialWidthDiff) complementaryGroups += 1;

    const sorted = [...group].sort((a, b) => scorePlistRecord(b) - scorePlistRecord(a));
    const keep = sorted[0];
    const remove = sorted.slice(1);

    duplicateGroups.push({
      code: codeKey,
      count: group.length,
      category: isExactDuplicate || onlyDescDiff
        ? 'exact'
        : hasConflict
          ? 'conflict'
          : onlyMaterialWidthDiff
            ? 'complementary'
            : 'mixed',
      keep: {
        rowIndex: keep.rowIndex,
        pqFormDesc: keep.pqFormDesc,
        type: keep.type,
        width: keep.width,
        height: keep.height,
        length: keep.length,
        thickness: keep.thickness,
        materialWidth: keep.materialWidth,
        score: scorePlistRecord(keep),
      },
      removeRows: remove.map((r) => ({
        rowIndex: r.rowIndex,
        pqFormDesc: r.pqFormDesc,
        type: r.type,
        width: r.width,
        height: r.height,
        length: r.length,
        thickness: r.thickness,
        materialWidth: r.materialWidth,
        score: scorePlistRecord(r),
      })),
      variants: sigEntries.map(([sig, rows]) => ({
        signature: sig,
        count: rows.length,
        sample: {
          rowIndex: rows[0].rowIndex,
          pqFormDesc: rows[0].pqFormDesc,
          type: rows[0].type,
          width: rows[0].width,
          height: rows[0].height,
          length: rows[0].length,
          thickness: rows[0].thickness,
          materialWidth: rows[0].materialWidth,
        },
        diffsFromKeep: diffRecords(keep, rows[0]),
      })),
    });
  }

  duplicateGroups.sort((a, b) => b.count - a.count);

  return {
    totalRows: records.length,
    uniqueCodes: byCode.size,
    duplicateCodeCount: duplicateGroups.length,
    duplicateRowCount: duplicateGroups.reduce((s, g) => s + g.count, 0),
    rowsToRemove: duplicateGroups.reduce((s, g) => s + g.removeRows.length, 0),
    categories: {
      exact: duplicateGroups.filter((g) => g.category === 'exact').length,
      conflict: duplicateGroups.filter((g) => g.category === 'conflict').length,
      complementary: duplicateGroups.filter((g) => g.category === 'complementary').length,
      mixed: duplicateGroups.filter((g) => g.category === 'mixed').length,
    },
    exactDuplicateRows,
    conflictingGroups,
    complementaryGroups,
    duplicateGroups,
  };
}

export function buildDedupePlan(analysis) {
  const deleteRowIndexes = [];
  const mergeUpdates = [];

  for (const group of analysis.duplicateGroups) {
    const keep = group.keep;
    let bestMaterialWidth = keep.materialWidth;

    for (const variant of group.variants) {
      if (variant.sample.materialWidth && !bestMaterialWidth) {
        bestMaterialWidth = variant.sample.materialWidth;
      }
    }

    if (!keep.materialWidth && bestMaterialWidth) {
      mergeUpdates.push({
        rowIndex: keep.rowIndex,
        materialWidth: bestMaterialWidth,
        code: group.code,
      });
    }

    for (const row of group.removeRows) {
      deleteRowIndexes.push(row.rowIndex);
    }
  }

  deleteRowIndexes.sort((a, b) => b - a);
  return { deleteRowIndexes, mergeUpdates };
}

function quoteSheetName(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

export async function applyPlistDedupe({ dryRun = true } = {}) {
  const {
    getMaterialWidthSheetsClient,
    loadPlistWithMaterialWidth,
    PLIST_SHEET_NAME,
    PQFORM_SPREADSHEET_ID,
  } = await import('./materialWidthSheets.js');

  const scopes = dryRun
    ? ['https://www.googleapis.com/auth/spreadsheets.readonly']
    : ['https://www.googleapis.com/auth/spreadsheets'];
  const { sheets } = getMaterialWidthSheetsClient(scopes);
  const plist = await loadPlistWithMaterialWidth(sheets);
  const analysis = analyzePlistDuplicates(plist.rows);
  const plan = buildDedupePlan(analysis);

  const result = {
    dryRun,
    totalRows: analysis.totalRows,
    uniqueCodes: analysis.uniqueCodes,
    duplicateCodeCount: analysis.duplicateCodeCount,
    duplicateRowCount: analysis.duplicateRowCount,
    rowsToRemove: analysis.rowsToRemove,
    categories: analysis.categories,
    duplicateGroups: analysis.duplicateGroups,
    plan: {
      deleteCount: plan.deleteRowIndexes.length,
      mergeCount: plan.mergeUpdates.length,
      deleteRowIndexes: plan.deleteRowIndexes,
      mergeUpdates: plan.mergeUpdates,
    },
  };

  if (dryRun) return result;

  for (const update of plan.mergeUpdates) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: PQFORM_SPREADSHEET_ID,
      range: `${quoteSheetName(PLIST_SHEET_NAME)}!I${update.rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[update.materialWidth]] },
    });
  }

  const sheetMeta = await sheets.spreadsheets.get({
    spreadsheetId: PQFORM_SPREADSHEET_ID,
    fields: 'sheets.properties',
  });
  const tab = (sheetMeta.data.sheets || []).find((s) => s.properties?.title === PLIST_SHEET_NAME);
  const sheetId = tab?.properties?.sheetId;
  if (sheetId == null) throw new Error('PQ-Form-plist tab not found');

  const deleteRequests = plan.deleteRowIndexes.map((rowIndex) => ({
    deleteDimension: {
      range: {
        sheetId,
        dimension: 'ROWS',
        startIndex: rowIndex - 1,
        endIndex: rowIndex,
      },
    },
  }));

  if (deleteRequests.length) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: PQFORM_SPREADSHEET_ID,
      requestBody: { requests: deleteRequests },
    });
  }

  return {
    ...result,
    applied: true,
    deleted: deleteRequests.length,
    merged: plan.mergeUpdates.length,
    rowsAfter: analysis.totalRows - deleteRequests.length,
  };
}
