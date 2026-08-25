import * as XLSX from "xlsx";

export type ParsedSeriesRow = {
  serie1: string;
  serie2: string;
  codigoSap: string | null;
  descripcion: string | null;
  cantidad: number;
};

const normalizeHeader = (header: string) =>
  header
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

const cleanSeries = (value: unknown): string => {
  if (value === null || value === undefined) return "";

  if (typeof value === "number") {
    return String(Math.trunc(value));
  }

  return String(value).trim();
};

const cleanText = (value: unknown): string | null => {
  if (value === null || value === undefined) return null;
  const text = String(value).trim();
  return text.length ? text : null;
};

export const parseExcelSeriesFile = async (file: File): Promise<ParsedSeriesRow[]> => {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });

  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("La planilla no contiene hojas.");
  }

  const sheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
    defval: "",
    raw: false,
  });

  const parsed: ParsedSeriesRow[] = [];

  for (const row of rows) {
    const normalizedRow: Record<string, unknown> = {};

    for (const [key, value] of Object.entries(row)) {
      normalizedRow[normalizeHeader(key)] = value;
    }

    const serie1 = cleanSeries(
      normalizedRow.serie1 ??
        normalizedRow.serie01 ??
        normalizedRow.serial1 ??
        normalizedRow.sn1 ??
        normalizedRow.imei1,
    );
    const serie2 = cleanSeries(
      normalizedRow.serie2 ??
        normalizedRow.serie02 ??
        normalizedRow.serial2 ??
        normalizedRow.sn2 ??
        normalizedRow.imei2,
    );

    if (!serie1 || !serie2) {
      continue;
    }

    const codigoSap = cleanText(
      normalizedRow.codigosap ??
        normalizedRow.sku ??
        normalizedRow.codigo ??
        normalizedRow.codigosku ??
        normalizedRow.codigosapsku,
    );
    const descripcion = cleanText(normalizedRow.descripcion ?? normalizedRow.description);

    const cantidadRaw = cleanText(normalizedRow.cantidad ?? normalizedRow.cantida);
    const cantidad = cantidadRaw ? Number.parseInt(cantidadRaw, 10) || 1 : 1;

    parsed.push({
      serie1,
      serie2,
      codigoSap,
      descripcion,
      cantidad,
    });
  }

  return parsed;
};

export const normalizeScanValue = (value: string) => value.trim();
