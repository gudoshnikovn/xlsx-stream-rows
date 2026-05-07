/** A single cell value. `null` represents an empty/missing cell. */
export type CellValue = string | number | boolean | Date | null;

/** A spreadsheet row. Length matches the highest column index seen in that row. */
export type Row = CellValue[];
