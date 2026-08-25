"use client";

import { parseExcelSeriesFile } from "@/lib/series";
import { useEffect, useMemo, useRef, useState, type CSSProperties } from "react";
import { List } from "react-window";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import ExcelJS from "exceljs";
import QRCode from "qrcode";
import * as XLSX from "xlsx";

type Association = {
  serie1: string;
  serie2: string;
  codigoSap: string | null;
  descripcion: string | null;
  cantidad: number;
};

type ScanStatus = "ok" | "duplicate" | "not_found";

type ScanEvent = {
  id: number;
  scannedSeries: string;
  serie1: string | null;
  serie2: string | null;
  codigoSap: string | null;
  descripcion: string | null;
  status: ScanStatus;
  message: string;
  createdAt: number;
};

type PalletSession = {
  id: number;
  lpn: string;
  targetQuantity: number;
  status: "active" | "finished";
  startedAt: number;
  finishedAt: number | null;
  sku: string;
  descripcion: string;
  okCount: number;
  duplicateCount: number;
  notFoundCount: number;
  scanEvents: ScanEvent[];
  readPairKeys: Set<string>;
};

type Notice = { type: "success" | "warning" | "error"; text: string };

const pairKey = (serie1: string, serie2: string) => `${serie1}||${serie2}`;

const statusClass: Record<ScanStatus, string> = {
  ok: "bg-emerald-500/15 text-emerald-100 border-emerald-400/40",
  duplicate: "bg-amber-500/15 text-amber-100 border-amber-400/40",
  not_found: "bg-rose-500/15 text-rose-100 border-rose-400/40",
};

const EventRow = ({ rows, onDelete, onView, currentSku, index, style }: any) => {
  const event = (rows as ScanEvent[])[index];
  const s1Read = event.scannedSeries === event.serie1;
  const s2Read = event.scannedSeries === event.serie2;
  const resolvedSku = event.codigoSap?.trim() || currentSku?.trim() || "SKU NO DISPONIBLE";

  return (
    <div
      style={style}
      className="group grid grid-cols-[100px_360px_360px_190px_160px_120px] items-center gap-x-4 border-b border-white/5 px-3 text-xs"
    >
      <span className={`w-fit rounded-full border px-2 py-0.5 ${statusClass[event.status]}`}>
        {event.status === "ok" ? "Válida" : event.status === "duplicate" ? "Repetida" : "No existe"}
      </span>
      <span
        title={event.serie1 ?? "-"}
        className={`font-mono whitespace-nowrap ${s1Read ? "rounded bg-sky-500/20 px-2 py-0.5 text-sky-100" : ""}`}
      >
        {event.serie1 ?? "-"}
      </span>
      <span
        title={event.serie2 ?? "-"}
        className={`font-mono whitespace-nowrap ${s2Read ? "rounded bg-sky-500/20 px-2 py-0.5 text-sky-100" : ""}`}
      >
        {event.serie2 ?? "-"}
      </span>
      <span title={resolvedSku} className={`font-mono whitespace-nowrap ${resolvedSku === "SKU NO DISPONIBLE" ? "text-amber-200" : "text-sky-100"}`}>
        {resolvedSku}
      </span>
      <span className="whitespace-nowrap">{new Date(event.createdAt).toLocaleTimeString("es-CL")}</span>
      <div className="flex justify-end gap-1 opacity-0 transition group-hover:opacity-100">
        <button
          type="button"
          onClick={() => onView(event)}
          className="rounded-md border border-sky-300/40 bg-sky-500/20 px-2 py-0.5 text-[11px] text-sky-100 hover:bg-sky-500/35"
          title="Ver detalle"
        >
          👁
        </button>
        <button
          type="button"
          onClick={() => onDelete(event.id)}
          className="rounded-md border border-rose-300/40 bg-rose-500/20 px-2 py-0.5 text-[11px] text-rose-100 hover:bg-rose-500/35"
          title="Borrar lectura"
        >
          🗑
        </button>
      </div>
    </div>
  );
};

const SessionRow = ({ rows, onSelect, onDelete, onExportExcel, onExportPdf, onQrCopy, activeId, index, style }: any) => {
  const session = (rows as PalletSession[])[index] as PalletSession;
  const isActive = session.id === activeId;

  return (
    <div
      style={style}
      onClick={() => onSelect(session)}
      className={`group grid cursor-pointer grid-cols-[150px_100px_120px_1fr_110px_410px] items-center border-b border-white/5 px-3 text-left text-xs hover:bg-white/5 ${isActive ? "bg-sky-500/15" : ""}`}
    >
      <span>{session.lpn}</span>
      <span>{session.targetQuantity}</span>
      <span>{session.sku || "-"}</span>
      <span className="truncate pr-2">{session.descripcion || "Sin lectura válida"}</span>
      <span className={session.status === "finished" ? "text-emerald-200" : "text-amber-200"}>
        {session.status === "finished" ? "Finalizado" : "Activo"}
      </span>

      <div className="flex justify-end gap-1 opacity-0 transition group-hover:opacity-100">
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            onSelect(session);
          }}
          className="rounded-md border border-sky-300/40 bg-sky-500/20 px-2 py-0.5 text-[11px] text-sky-100 hover:bg-sky-500/35"
          title="Ver series"
        >
          Ver series
        </button>
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            onExportExcel(session.id);
          }}
          className="rounded-md border border-emerald-300/40 bg-emerald-500/20 px-2 py-0.5 text-[11px] text-emerald-100 hover:bg-emerald-500/35"
          title="Exportar línea a Excel"
        >
          Excel
        </button>
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            onExportPdf(session.id);
          }}
          className="rounded-md border border-violet-300/40 bg-violet-500/20 px-2 py-0.5 text-[11px] text-violet-100 hover:bg-violet-500/35"
          title="Exportar línea a PDF"
        >
          PDF
        </button>
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            onQrCopy(session.id);
          }}
          className="rounded-md border border-cyan-300/40 bg-cyan-500/20 px-2 py-0.5 text-[11px] text-cyan-100 hover:bg-cyan-500/35"
          title="Excel con QR y copia"
        >
          Qr/Copiar
        </button>
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            onDelete(session.id);
          }}
          className="rounded-md border border-rose-300/40 bg-rose-500/20 px-2 py-0.5 text-[11px] text-rose-100 hover:bg-rose-500/35"
          title="Borrar pallet"
        >
          Borrar
        </button>
      </div>
    </div>
  );
};

export function LocalSeriesPlatform() {
  const [associationsMap, setAssociationsMap] = useState<Map<string, Association>>(new Map());
  const [importInfo, setImportInfo] = useState<{ fileName: string; rowCount: number } | null>(null);
  const [importDefaultSku, setImportDefaultSku] = useState("");

  const [sessions, setSessions] = useState<PalletSession[]>([]);
  const [activeSessionId, setActiveSessionId] = useState<number | null>(null);

  const [lpn, setLpn] = useState("");
  const [packagingTarget, setPackagingTarget] = useState("1");
  const [scanInput, setScanInput] = useState("");

  const [isImporting, setIsImporting] = useState(false);
  const [isPastePreviewing, setIsPastePreviewing] = useState(false);
  const [notice, setNotice] = useState<Notice | null>(null);
  const [finishToast, setFinishToast] = useState<{ lpn: string; ok: number; target: number } | null>(null);
  const [invalidToast, setInvalidToast] = useState<{ value: string } | null>(null);
  const [selectedRead, setSelectedRead] = useState<ScanEvent | null>(null);
  const [showCapturePalletModal, setShowCapturePalletModal] = useState(false);
  const [nextPalletLpn, setNextPalletLpn] = useState("");
  const [nextPalletTarget, setNextPalletTarget] = useState("1");

  const fileRef = useRef<HTMLInputElement | null>(null);
  const scanRef = useRef<HTMLInputElement | null>(null);
  const readingsSectionRef = useRef<HTMLElement | null>(null);
  const pasteTimerRef = useRef<number | null>(null);
  const scanTimerRef = useRef<number | null>(null);
  const eventIdRef = useRef(1);
  const sessionIdRef = useRef(1);

  const activeSession = useMemo(
    () => sessions.find((item) => item.id === activeSessionId) ?? null,
    [sessions, activeSessionId],
  );

  const handleImport = async (file: File) => {
    setIsImporting(true);
    setNotice(null);

    try {
      const parsed = await parseExcelSeriesFile(file);

      if (!parsed.length) {
        throw new Error(
          "No se detectaron filas válidas. Verifica columnas Serie 1 y Serie 2 en la planilla.",
        );
      }

      const map = new Map<string, Association>();
      const firstDetectedSku = parsed.find((row) => row.codigoSap?.trim())?.codigoSap?.trim() ?? "";

      for (const row of parsed) {
        const resolvedSku = row.codigoSap?.trim() || firstDetectedSku || null;

        const assoc: Association = {
          serie1: row.serie1,
          serie2: row.serie2,
          codigoSap: resolvedSku,
          descripcion: row.descripcion,
          cantidad: row.cantidad,
        };

        map.set(row.serie1, assoc);
        map.set(row.serie2, assoc);
      }

      setAssociationsMap(map);
      setImportDefaultSku(firstDetectedSku);
      setImportInfo({ fileName: file.name, rowCount: parsed.length });
      setSessions([]);
      setActiveSessionId(null);
      setNotice({ type: "success", text: `Planilla cargada (${parsed.length.toLocaleString("es-CL")} filas).` });
    } catch (error) {
      const message = error instanceof Error ? error.message : "No se pudo importar el Excel.";
      setNotice({ type: "error", text: message });
    } finally {
      setIsImporting(false);
      if (fileRef.current) fileRef.current.value = "";
    }
  };

  const createPalletSession = (rawLpn: string, rawTarget: string) => {
    if (!associationsMap.size) {
      setNotice({ type: "error", text: "Debes importar una planilla primero." });
      return false;
    }

    if (sessions.length >= 100) {
      setNotice({ type: "error", text: "Se alcanzó el límite máximo de 100 pallets." });
      return false;
    }

    const normalizedLpn = rawLpn.trim();
    const target = Number.parseInt(rawTarget, 10);

    if (!normalizedLpn || !Number.isInteger(target) || target <= 0) {
      setNotice({ type: "error", text: "Completa LPN y Paquetería pallet (>0)." });
      return false;
    }

    const lpnExists = sessions.some((session) => session.lpn.toLowerCase() === normalizedLpn.toLowerCase());
    if (lpnExists) {
      setNotice({ type: "error", text: `El LPN ${normalizedLpn} ya existe. Ingresa uno distinto.` });
      return false;
    }

    const next: PalletSession = {
      id: sessionIdRef.current++,
      lpn: normalizedLpn,
      targetQuantity: target,
      status: "active",
      startedAt: Date.now(),
      finishedAt: null,
      sku: importDefaultSku,
      descripcion: "",
      okCount: 0,
      duplicateCount: 0,
      notFoundCount: 0,
      scanEvents: [],
      readPairKeys: new Set<string>(),
    };

    setSessions((prev) => [next, ...prev]);
    setActiveSessionId(next.id);
    setScanInput("");
    setNotice({ type: "success", text: `Pallet ${next.lpn} iniciado.` });
    setTimeout(() => scanRef.current?.focus(), 40);
    return true;
  };

  const startPallet = () => {
    void createPalletSession(lpn, packagingTarget);
  };

  const pushScan = (value: string) => {
    if (!activeSessionId) return;

    if (activeSession?.status !== "active") {
      if (activeSession) {
        setNotice({ type: "warning", text: `⚠️ Paquetería culminada para pallet ${activeSession.lpn}.` });
        setFinishToast({
          lpn: activeSession.lpn,
          ok: activeSession.okCount,
          target: activeSession.targetQuantity,
        });
      }
      setScanInput("");
      return;
    }

    const assoc = associationsMap.get(value);
    if (!assoc) {
      setNotice({ type: "error", text: "Serie no corresponde. Favor leer la serie correcta." });
      setInvalidToast({ value });
      setScanInput("");
      return;
    }

    setSessions((prev) => {
      const idx = prev.findIndex((s) => s.id === activeSessionId);
      if (idx < 0) return prev;

      const session = prev[idx];
      if (session.status !== "active") return prev;

      const now = Date.now();
      let event: ScanEvent;
      let updated: PalletSession;

      const key = pairKey(assoc.serie1, assoc.serie2);

        if (session.readPairKeys.has(key)) {
          const effectiveSku = assoc.codigoSap?.trim() || session.sku?.trim() || importDefaultSku;

          event = {
            id: eventIdRef.current++,
            scannedSeries: value,
            serie1: assoc.serie1,
            serie2: assoc.serie2,
            codigoSap: effectiveSku,
            descripcion: assoc.descripcion,
            status: "duplicate",
            message: "Serie repetida: esa pareja ya fue leída.",
            createdAt: now,
          };

          updated = {
            ...session,
            duplicateCount: session.duplicateCount + 1,
            scanEvents: [event, ...session.scanEvents],
          };

          setNotice({ type: "warning", text: event.message });
        } else {
          const effectiveSku = assoc.codigoSap?.trim() || session.sku?.trim() || importDefaultSku;

          event = {
            id: eventIdRef.current++,
            scannedSeries: value,
            serie1: assoc.serie1,
            serie2: assoc.serie2,
            codigoSap: effectiveSku,
            descripcion: assoc.descripcion,
            status: "ok",
            message: "Lectura válida.",
            createdAt: now,
          };

          const readPairKeys = new Set(session.readPairKeys);
          readPairKeys.add(key);

          const okCount = session.okCount + 1;
          const finished = okCount >= session.targetQuantity;

          const nextSku = effectiveSku || session.sku || "";

          updated = {
            ...session,
            sku: nextSku,
            descripcion: assoc.descripcion?.trim() || session.descripcion,
            okCount,
            status: finished ? "finished" : session.status,
            finishedAt: finished ? now : session.finishedAt,
            readPairKeys,
            scanEvents: [event, ...session.scanEvents],
          };

          if (finished) {
            setNotice({
              type: "success",
              text: `✅ Finalizada la paquetería del pallet ${session.lpn}. Lecturas válidas: ${okCount}/${session.targetQuantity}.`,
            });
            setFinishToast({ lpn: session.lpn, ok: okCount, target: session.targetQuantity });
          } else {
            setNotice({ type: "success", text: `OK: ${assoc.serie1} ↔ ${assoc.serie2}` });
          }
        }

      const next = [...prev];
      next[idx] = updated;
      return next;
    });

    setScanInput("");
  };

  const downloadWorkbook = (wb: XLSX.WorkBook, fileName: string) => {
    try {
      XLSX.writeFileXLSX(wb, fileName);
      return;
    } catch {
      const bytes = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([bytes], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1500);
    }
  };

  const deleteScanEvent = (eventId: number) => {
    if (!activeSessionId) return;

    setSessions((prev) =>
      prev.map((session) => {
        if (session.id !== activeSessionId) return session;

        const nextEvents = session.scanEvents.filter((event) => event.id !== eventId);
        if (nextEvents.length === session.scanEvents.length) return session;

        const okEvents = nextEvents.filter((event) => event.status === "ok");
        const duplicateCount = nextEvents.filter((event) => event.status === "duplicate").length;
        const notFoundCount = nextEvents.filter((event) => event.status === "not_found").length;

        const readPairKeys = new Set<string>();
        for (const event of okEvents) {
          if (event.serie1 && event.serie2) {
            readPairKeys.add(pairKey(event.serie1, event.serie2));
          }
        }

        const latestOk = okEvents[0];
        const okCount = okEvents.length;
        const isFinished = okCount >= session.targetQuantity;

        return {
          ...session,
          scanEvents: nextEvents,
          readPairKeys,
          okCount,
          duplicateCount,
          notFoundCount,
          sku: latestOk?.codigoSap?.trim() || importDefaultSku,
          descripcion: latestOk?.descripcion?.trim() || "",
          status: isFinished ? "finished" : "active",
          finishedAt: isFinished ? session.finishedAt ?? Date.now() : null,
        };
      }),
    );

    setNotice({ type: "success", text: "Lectura eliminada correctamente." });
  };

  const deletePalletSession = (sessionId: number) => {
    setSessions((prev) => prev.filter((session) => session.id !== sessionId));

    if (activeSessionId === sessionId) {
      setActiveSessionId(null);
      setScanInput("");
    }

    setNotice({ type: "success", text: "Pallet eliminado correctamente." });
  };

  const viewPalletSession = (targetSession: PalletSession) => {
    setActiveSessionId(targetSession.id);
    setSelectedRead(null);
    setScanInput("");
    setNotice({ type: "success", text: `Mostrando series del pallet ${targetSession.lpn}.` });

    window.requestAnimationFrame(() => {
      readingsSectionRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  };

  const exportPallet = (session: PalletSession) => {
    try {
      const valid = session.scanEvents.filter((event) => event.status === "ok");
      const rows = valid.length
        ? valid.map((event) => ({
            LPN: session.lpn,
            SKU: event.codigoSap?.trim() || session.sku,
            DESCRIPCION: event.descripcion ?? "",
            PAQUETERIA_PALLET: session.targetQuantity,
            SERIES_1: event.serie1,
            SERIES_2: event.serie2,
            SERIE_LEIDA: event.scannedSeries,
            FECHA_LECTURA: new Date(event.createdAt).toISOString(),
          }))
        : [
            {
              LPN: session.lpn,
              SKU: session.sku,
              DESCRIPCION: session.descripcion || "Sin lecturas válidas",
              PAQUETERIA_PALLET: session.targetQuantity,
              SERIES_1: "",
              SERIES_2: "",
              SERIE_LEIDA: "",
              FECHA_LECTURA: "",
            },
          ];

      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Pallet");
      downloadWorkbook(wb, `pallet_${session.lpn}_${session.id}.xlsx`);
      setNotice({ type: "success", text: `Excel generado para pallet ${session.lpn}.` });
    } catch {
      setNotice({ type: "error", text: "No se pudo generar el Excel del pallet." });
    }
  };

  const exportCurrent = () => {
    const session = activeSession ?? sessions[0];
    if (!session) {
      setNotice({ type: "error", text: "No hay pallets para exportar." });
      return;
    }
    exportPallet(session);
  };

  const exportPalletPdf = (session: PalletSession) => {
    try {
      const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
      const valid = session.scanEvents.filter((event) => event.status === "ok");

      doc.setFontSize(14);
      doc.text(`GESTOR DE SERIES - PALLET ${session.lpn}`, 14, 12);
      doc.setFontSize(10);
      doc.text(`Paquetería: ${session.targetQuantity} | Estado: ${session.status}`, 14, 18);

      const body = (valid.length
        ? valid
        : [
            {
              serie1: "",
              serie2: "",
              codigoSap: session.sku,
              descripcion: session.descripcion || "Sin lecturas válidas",
              scannedSeries: "",
              createdAt: Date.now(),
            },
          ]).map((event) => [
        event.serie1 || "",
        event.serie2 || "",
        event.codigoSap || "",
        event.descripcion || "",
        event.scannedSeries || "",
        event.createdAt ? new Date(event.createdAt).toLocaleString("es-CL") : "",
      ]);

      autoTable(doc, {
        startY: 24,
        head: [["Series 1", "Series 2", "SKU", "Descripción", "Serie leída", "Fecha"]],
        body,
        theme: "grid",
        styles: { fontSize: 8 },
        headStyles: { fillColor: [30, 41, 59] },
      });

      doc.save(`pallet_${session.lpn}_${session.id}.pdf`);
      setNotice({ type: "success", text: `PDF generado para pallet ${session.lpn}.` });
    } catch {
      setNotice({ type: "error", text: "No se pudo generar el PDF del pallet." });
    }
  };

  const exportCurrentPdf = () => {
    const session = activeSession ?? sessions[0];
    if (!session) {
      setNotice({ type: "error", text: "No hay pallets para exportar en PDF." });
      return;
    }

    exportPalletPdf(session);
  };

  const exportLineExcel = (sessionId: number) => {
    const session = sessions.find((item) => item.id === sessionId);
    if (!session) {
      setNotice({ type: "error", text: "No se encontró el pallet para exportar." });
      return;
    }
    exportPallet(session);
  };

  const exportLinePdf = (sessionId: number) => {
    const session = sessions.find((item) => item.id === sessionId);
    if (!session) {
      setNotice({ type: "error", text: "No se encontró el pallet para exportar." });
      return;
    }
    exportPalletPdf(session);
  };

  const exportLineQrCopy = async (sessionId: number) => {
    const session = sessions.find((item) => item.id === sessionId);
    if (!session) {
      setNotice({ type: "error", text: "No se encontró el pallet para exportar." });
      return;
    }

    try {
      const valid = session.scanEvents.filter((event) => event.status === "ok");
      const series1 = valid.map((event) => (event.serie1 ?? "").trim()).filter(Boolean);
      const series2 = valid.map((event) => (event.serie2 ?? "").trim()).filter(Boolean);

      const series1Joined = series1.join(";");
      const series2Joined = series2.join(";");

      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet("QR-Copiar");

      sheet.columns = [{ width: 42 }, { width: 42 }];

      sheet.getCell("A1").value = "SERIES_1";
      sheet.getCell("B1").value = "SERIES_2";
      sheet.getCell("A2").value = "Agrupación de series para copia pega";
      sheet.getCell("B2").value = "Agrupación de series para copia pega";
      sheet.getCell("A3").value = series1Joined;
      sheet.getCell("B3").value = series2Joined;
      sheet.getCell("A8").value = "Qr series 1";
      sheet.getCell("B8").value = "Qr series 2";

      for (const cellRef of ["A1", "B1", "A2", "B2", "A8", "B8"]) {
        const cell = sheet.getCell(cellRef);
        cell.font = { bold: true, size: 11 };
        cell.alignment = { vertical: "middle", horizontal: "center" };
      }

      sheet.getCell("A3").alignment = { wrapText: true, vertical: "top", horizontal: "left" };
      sheet.getCell("B3").alignment = { wrapText: true, vertical: "top", horizontal: "left" };

      sheet.getRow(1).height = 22;
      sheet.getRow(2).height = 20;
      sheet.getRow(3).height = 80;
      sheet.getRow(8).height = 22;

      const addBorderRange = (startCol: string, endCol: string, startRow: number, endRow: number) => {
        for (let row = startRow; row <= endRow; row += 1) {
          for (let colCode = startCol.charCodeAt(0); colCode <= endCol.charCodeAt(0); colCode += 1) {
            const cell = sheet.getCell(`${String.fromCharCode(colCode)}${row}`);
            cell.border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }
        }
      };

      addBorderRange("A", "B", 1, 12);

      const qr1Text = series1Joined || `Pallet ${session.lpn} sin series 1`;
      const qr2Text = series2Joined || `Pallet ${session.lpn} sin series 2`;

      const qr1 = await QRCode.toDataURL(qr1Text, { margin: 1, width: 260 });
      const qr2 = await QRCode.toDataURL(qr2Text, { margin: 1, width: 260 });

      const qr1Id = workbook.addImage({ base64: qr1, extension: "png" });
      const qr2Id = workbook.addImage({ base64: qr2, extension: "png" });

      sheet.addImage(qr1Id, { tl: { col: 0.15, row: 8.2 }, ext: { width: 180, height: 180 } });
      sheet.addImage(qr2Id, { tl: { col: 1.15, row: 8.2 }, ext: { width: 180, height: 180 } });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `qr_copiar_${session.lpn}_${session.id}.xlsx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1500);

      setNotice({ type: "success", text: `Excel QR/Copiar generado para pallet ${session.lpn}.` });
    } catch {
      setNotice({ type: "error", text: "No se pudo generar el Excel QR/Copiar." });
    }
  };

  const exportAll = () => {
    if (!sessions.length) {
      setNotice({ type: "error", text: "No hay pallets registrados." });
      return;
    }

    try {
      const rows: Array<Record<string, unknown>> = [];

      for (const session of sessions) {
        const valid = session.scanEvents.filter((event) => event.status === "ok");
        if (!valid.length) {
          rows.push({
            LPN: session.lpn,
            SKU: session.sku,
            DESCRIPCION: session.descripcion || "Sin lecturas válidas",
            ESTADO: session.status,
            PAQUETERIA_PALLET: session.targetQuantity,
            SERIES_1: "",
            SERIES_2: "",
            SERIE_LEIDA: "",
          });
          continue;
        }

        for (const event of valid) {
          rows.push({
            LPN: session.lpn,
            SKU: event.codigoSap?.trim() || session.sku,
            DESCRIPCION: event.descripcion ?? "",
            ESTADO: session.status,
            PAQUETERIA_PALLET: session.targetQuantity,
            SERIES_1: event.serie1,
            SERIES_2: event.serie2,
            SERIE_LEIDA: event.scannedSeries,
          });
        }
      }

      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Pallets");
      downloadWorkbook(wb, "pallets_consolidados.xlsx");
      setNotice({ type: "success", text: "Excel consolidado generado correctamente." });
    } catch {
      setNotice({ type: "error", text: "No se pudo generar el Excel consolidado." });
    }
  };

  const exportAllPdf = () => {
    if (!sessions.length) {
      setNotice({ type: "error", text: "No hay pallets registrados para PDF." });
      return;
    }

    try {
      const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
      doc.setFontSize(14);
      doc.text("GESTOR DE SERIES - CONSOLIDADO DE PALLETS", 14, 12);

      const rows: Array<Array<string>> = [];
      for (const session of sessions) {
        const valid = session.scanEvents.filter((event) => event.status === "ok");

        if (!valid.length) {
          rows.push([
            session.lpn,
            String(session.targetQuantity),
            session.sku || "",
            session.descripcion || "Sin lecturas válidas",
            session.status,
            "",
            "",
            "",
          ]);
          continue;
        }

        for (const event of valid) {
          rows.push([
            session.lpn,
            String(session.targetQuantity),
            event.codigoSap?.trim() || session.sku || "",
            event.descripcion || "",
            session.status,
            event.serie1 || "",
            event.serie2 || "",
            event.scannedSeries || "",
          ]);
        }
      }

      autoTable(doc, {
        startY: 18,
        head: [["LPN", "Paq", "SKU", "Descripción", "Estado", "Series 1", "Series 2", "Serie leída"]],
        body: rows,
        theme: "grid",
        styles: { fontSize: 7 },
        headStyles: { fillColor: [30, 41, 59] },
      });

      doc.save("pallets_consolidados.pdf");
      setNotice({ type: "success", text: "PDF consolidado generado correctamente." });
    } catch {
      setNotice({ type: "error", text: "No se pudo generar el PDF consolidado." });
    }
  };

  const resetAll = () => {
    setAssociationsMap(new Map());
    setImportDefaultSku("");
    setImportInfo(null);
    setSessions([]);
    setActiveSessionId(null);
    setLpn("");
    setPackagingTarget("1");
    setNextPalletLpn("");
    setNextPalletTarget("1");
    setShowCapturePalletModal(false);
    setScanInput("");
    setNotice(null);
  };

  const captureNextPallet = () => {
    const created = createPalletSession(nextPalletLpn, nextPalletTarget);
    if (!created) return;

    setNextPalletLpn("");
    setNextPalletTarget("1");
    setShowCapturePalletModal(false);
  };

  useEffect(() => {
    if (!activeSession || activeSession.status !== "active" || !scanInput.trim() || isPastePreviewing) {
      return;
    }

    if (scanTimerRef.current) window.clearTimeout(scanTimerRef.current);
    scanTimerRef.current = window.setTimeout(() => pushScan(scanInput.trim()), 180);

    return () => {
      if (scanTimerRef.current) window.clearTimeout(scanTimerRef.current);
    };
  }, [scanInput, activeSession, isPastePreviewing]);

  useEffect(() => {
    return () => {
      if (pasteTimerRef.current) window.clearTimeout(pasteTimerRef.current);
      if (scanTimerRef.current) window.clearTimeout(scanTimerRef.current);
    };
  }, []);

  useEffect(() => {
    if (!finishToast) return;
    const timer = window.setTimeout(() => setFinishToast(null), 3200);
    return () => window.clearTimeout(timer);
  }, [finishToast]);

  useEffect(() => {
    if (!invalidToast) return;
    const timer = window.setTimeout(() => setInvalidToast(null), 2200);
    return () => window.clearTimeout(timer);
  }, [invalidToast]);

  const rowsForHistory = activeSession?.scanEvents ?? [];

  return (
    <main className="min-h-screen bg-[radial-gradient(circle_at_0%_0%,#1d4ed8_0%,transparent_45%),radial-gradient(circle_at_100%_0%,#0f766e_0%,transparent_35%),linear-gradient(135deg,#020617_0%,#0f172a_45%,#111827_100%)] p-4 pb-24 text-slate-100 md:p-6">
      {finishToast && (
        <div className="fixed right-4 top-4 z-50 w-[min(92vw,430px)] rounded-2xl border border-emerald-200/40 bg-emerald-500/90 px-4 py-3 text-white shadow-[0_16px_36px_rgba(16,185,129,0.35)] backdrop-blur">
          <p className="text-sm font-semibold">✅ Paquetería culminada</p>
          <p className="mt-1 text-xs">
            Pallet <strong>{finishToast.lpn}</strong> finalizado con {finishToast.ok}/{finishToast.target} lecturas válidas.
          </p>
        </div>
      )}

      {invalidToast && (
        <div className="fixed right-4 top-24 z-50 w-[min(92vw,430px)] rounded-2xl border border-rose-200/40 bg-rose-500/90 px-4 py-3 text-white shadow-[0_16px_36px_rgba(244,63,94,0.35)] backdrop-blur">
          <p className="text-sm font-semibold">⚠️ Serie no corresponde</p>
          <p className="mt-1 text-xs">
            La serie <strong>{invalidToast.value}</strong> no coincide con la planilla y no fue cargada.
          </p>
        </div>
      )}

      <div className="mx-auto grid w-full max-w-7xl gap-4">
        <header className="rounded-3xl border border-white/15 bg-white/8 p-5 text-center shadow-[0_20px_50px_rgba(2,8,23,0.45)] backdrop-blur-xl">
          <h1 className="text-2xl font-bold tracking-tight text-white md:text-4xl">GESTOR DE SERIES</h1>
          <p className="mt-1 text-xs text-slate-300 md:text-sm">Plataforma local · rápida · optimizada para alto volumen</p>
        </header>

        <section className="grid gap-3 md:grid-cols-[1.2fr_1fr]">
          <article className="rounded-3xl border border-white/10 bg-white/6 p-4 shadow-[0_16px_40px_rgba(15,23,42,0.35)] backdrop-blur-lg">
            <h2 className="text-sm font-semibold text-white">1) Importación Excel</h2>
            <div className="mt-2 flex flex-wrap items-center gap-2">
              <input
                ref={fileRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (file) void handleImport(file);
                }}
                className="rounded-xl border border-white/20 bg-slate-900/70 px-3 py-1.5 text-xs shadow-inner outline-none ring-sky-400/60 transition focus:ring"
              />
              {isImporting && <span className="text-[11px] text-sky-200">Importando...</span>}
            </div>
            <p className="mt-3 text-xs text-slate-300">
              {importInfo
                ? `Archivo: ${importInfo.fileName} · Filas: ${importInfo.rowCount.toLocaleString("es-CL")} · SKU detectado: ${importDefaultSku || "No disponible"}`
                : "Aún no hay planilla cargada."}
            </p>
          </article>

          <article className="rounded-3xl border border-white/10 bg-white/6 p-4 shadow-[0_16px_40px_rgba(15,23,42,0.35)] backdrop-blur-lg">
            <h2 className="text-sm font-semibold text-white">2) Configurar pallet</h2>
            <div className="mt-2 grid gap-1.5">
              <input value={lpn} onChange={(e) => setLpn(e.target.value)} placeholder="LPN" className="rounded-xl border border-white/20 bg-slate-900/70 px-3 py-1.5 text-xs shadow-inner outline-none ring-sky-400/60 transition focus:ring" />
              <label className="text-[11px] text-slate-300">Paquetería pallet</label>
              <input type="number" min={1} value={packagingTarget} onChange={(e) => setPackagingTarget(e.target.value)} placeholder="Ingresa cantidad" className="rounded-xl border border-white/20 bg-slate-900/70 px-3 py-1.5 text-xs shadow-inner outline-none ring-sky-400/60 transition focus:ring" />
              <button type="button" onClick={startPallet} className="rounded-xl bg-gradient-to-r from-sky-500 to-cyan-500 px-3 py-1.5 text-xs font-semibold text-white shadow hover:from-sky-400 hover:to-cyan-400">
                Iniciar lectura
              </button>
            </div>
          </article>
        </section>

        {notice && (
          <div className={`rounded-2xl border px-3 py-2 text-sm shadow-lg backdrop-blur ${notice.type === "success" ? "border-emerald-300/40 bg-emerald-500/15 text-emerald-100" : notice.type === "warning" ? "border-amber-300/40 bg-amber-500/15 text-amber-100" : "border-rose-300/40 bg-rose-500/15 text-rose-100"}`}>
            {notice.text}
          </div>
        )}

        <section className="grid gap-4 lg:grid-cols-[1fr_1.25fr]">
          <article className="rounded-3xl border border-white/10 bg-white/6 p-5 shadow-[0_16px_40px_rgba(15,23,42,0.35)] backdrop-blur-lg">
            <h2 className="font-semibold text-white">3) Lectura</h2>
            <div className="mt-3 grid gap-2">
              <input
                ref={scanRef}
                value={scanInput}
                onChange={(e) => setScanInput(e.target.value)}
                onPaste={(e) => {
                  e.preventDefault();
                  const pasted = e.clipboardData.getData("text").trim();
                  setScanInput(pasted);
                  if (!pasted) return;

                  setIsPastePreviewing(true);
                  if (pasteTimerRef.current) window.clearTimeout(pasteTimerRef.current);
                  pasteTimerRef.current = window.setTimeout(() => {
                    setIsPastePreviewing(false);
                    pushScan(pasted);
                  }, 300);
                }}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    e.preventDefault();
                    pushScan(scanInput.trim());
                  }
                }}
                disabled={!activeSession || activeSession.status !== "active"}
                placeholder="Escanea o pega una serie"
                className="rounded-xl border border-white/20 bg-slate-900/80 px-3 py-3 text-lg shadow-inner outline-none ring-sky-400/60 transition focus:ring"
              />

              {activeSession?.status === "finished" && (
                <div className="rounded-xl border border-emerald-300/50 bg-emerald-500/20 px-3 py-2 text-xs text-emerald-100">
                  ✅ Paquetería culminada. No se permiten más lecturas para este pallet.
                </div>
              )}

              <div className="rounded-xl border border-sky-300/30 bg-sky-500/15 px-3 py-2 text-center text-xs text-sky-100">
                Leídas vs Paquetería: <strong>{activeSession?.okCount ?? 0}</strong> / <strong>{activeSession?.targetQuantity ?? 0}</strong>
              </div>

              <div className="grid grid-cols-3 gap-2 text-center text-xs">
                <div className="rounded-xl border border-emerald-300/30 bg-emerald-500/15 p-2">Válidas: {activeSession?.okCount ?? 0}</div>
                <div className="rounded-xl border border-amber-300/30 bg-amber-500/15 p-2">Repetidas: {activeSession?.duplicateCount ?? 0}</div>
                <div className="rounded-xl border border-rose-300/30 bg-rose-500/15 p-2">No existe: {activeSession?.notFoundCount ?? 0}</div>
              </div>

              <div className="flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={() => setShowCapturePalletModal(true)}
                  disabled={!activeSession || activeSession.status !== "finished" || sessions.length >= 100}
                  className="rounded-xl border border-cyan-300/40 bg-cyan-500/20 px-3 py-2 text-xs font-medium hover:bg-cyan-500/30 disabled:opacity-50"
                >
                  Agregar otro pallet
                </button>
              </div>

            </div>
          </article>

          <article
            ref={readingsSectionRef}
            className="rounded-3xl border border-white/10 bg-white/6 p-5 shadow-[0_16px_40px_rgba(15,23,42,0.35)] backdrop-blur-lg"
          >
            <h2 className="font-semibold text-white">Lecturas (virtualizado)</h2>
            <div className="mt-3 overflow-x-auto rounded-t-xl border border-white/10 border-b-0 bg-white/10">
              <div className="grid min-w-[1320px] grid-cols-[100px_360px_360px_190px_160px_120px] gap-x-4 px-3 py-2 text-xs font-semibold text-slate-200">
                <span>Estado</span>
                <span>Series 1</span>
                <span>Series 2</span>
                <span>SKU</span>
                <span>Hora</span>
                <span className="text-right">Acción</span>
              </div>
            </div>
            <div className="overflow-x-auto rounded-b-xl border border-white/10 border-t-0 bg-slate-950/25">
              <div className="min-w-[1320px]">
                <List
                  rowComponent={EventRow}
                  rowCount={rowsForHistory.length}
                  rowHeight={40}
                  rowProps={{
                    rows: rowsForHistory,
                    onDelete: deleteScanEvent,
                    onView: setSelectedRead,
                    currentSku: activeSession?.sku ?? "",
                  }}
                  overscanCount={8}
                  style={{ width: 1320, height: 320 }}
                />
              </div>
            </div>
          </article>
        </section>

        <section className="rounded-3xl border border-white/10 bg-white/6 p-5 shadow-[0_16px_40px_rgba(15,23,42,0.35)] backdrop-blur-lg">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <h2 className="font-semibold text-white">Consultas de pallets registrados (virtualizado)</h2>
            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={exportAll}
                disabled={!sessions.length}
                className="rounded-xl border border-sky-300/40 bg-sky-500/20 px-3 py-2 text-xs font-medium hover:bg-sky-500/30 disabled:opacity-50"
              >
                Exportar todo a Excel
              </button>
              <button
                type="button"
                onClick={exportAllPdf}
                disabled={!sessions.length}
                className="rounded-xl border border-violet-300/40 bg-violet-500/20 px-3 py-2 text-xs font-medium hover:bg-violet-500/30 disabled:opacity-50"
              >
                Exportar todo a PDF
              </button>
            </div>
          </div>
          <div className="mt-3 grid grid-cols-[150px_100px_120px_1fr_110px_410px] rounded-t-xl bg-white/10 px-3 py-2 text-xs font-semibold text-slate-200">
            <span>LPN</span>
            <span>Paq.</span>
            <span>SKU</span>
            <span>Descripción</span>
            <span>Estado</span>
            <span className="text-right">Acción</span>
          </div>
          <div className="overflow-hidden rounded-b-xl border border-white/10 bg-slate-950/25">
            <List
              rowComponent={SessionRow}
              rowCount={sessions.length}
              rowHeight={38}
              rowProps={{
                rows: sessions,
                onSelect: viewPalletSession,
                onDelete: deletePalletSession,
                onExportExcel: exportLineExcel,
                onExportPdf: exportLinePdf,
                onQrCopy: exportLineQrCopy,
                activeId: activeSessionId,
              }}
              overscanCount={10}
              style={{ width: "100%", height: 260 }}
            />
          </div>
        </section>
      </div>

      <button
        type="button"
        onClick={resetAll}
        className="fixed right-5 top-20 z-40 rounded-full border border-rose-300/45 bg-rose-500/85 px-4 py-2 text-xs font-semibold text-white shadow-xl backdrop-blur hover:bg-rose-500"
      >
        Iniciar nueva tarea
      </button>

      {selectedRead && (
        <div className="fixed inset-0 z-50 grid place-items-center bg-slate-950/70 p-4">
          <div className="w-full max-w-lg rounded-2xl border border-sky-200/30 bg-slate-900/95 p-5 shadow-2xl backdrop-blur">
            <h3 className="text-base font-semibold text-white">Detalle de lectura</h3>
            <div className="mt-3 grid gap-2 text-xs text-slate-200">
              <p><span className="text-slate-400">Estado:</span> {selectedRead.status}</p>
              <p><span className="text-slate-400">Series 1:</span> {selectedRead.serie1 ?? "-"}</p>
              <p><span className="text-slate-400">Series 2:</span> {selectedRead.serie2 ?? "-"}</p>
              <p><span className="text-slate-400">SKU:</span> {selectedRead.codigoSap?.trim() || activeSession?.sku || "SKU NO DISPONIBLE"}</p>
              <p><span className="text-slate-400">Serie leída:</span> {selectedRead.scannedSeries}</p>
              <p><span className="text-slate-400">Hora:</span> {new Date(selectedRead.createdAt).toLocaleString("es-CL")}</p>
            </div>

            <div className="mt-4 flex justify-end">
              <button
                type="button"
                onClick={() => setSelectedRead(null)}
                className="rounded-xl border border-white/20 bg-white/10 px-3 py-2 text-xs font-medium"
              >
                Cerrar
              </button>
            </div>
          </div>
        </div>
      )}

      {showCapturePalletModal && (
        <div className="fixed inset-0 z-50 grid place-items-center bg-slate-950/70 p-4">
          <div className="w-full max-w-md rounded-2xl border border-cyan-200/30 bg-slate-900/95 p-5 shadow-2xl backdrop-blur">
            <h3 className="text-base font-semibold text-white">Capturar nuevo pallet</h3>
            <p className="mt-1 text-xs text-slate-300">Ingresa nuevo LPN (sin repetir) y paquetería para iniciar lectura.</p>

            <div className="mt-4 grid gap-2">
              <input
                value={nextPalletLpn}
                onChange={(e) => setNextPalletLpn(e.target.value)}
                placeholder="Nuevo LPN"
                className="rounded-xl border border-white/20 bg-slate-900/70 px-3 py-2 text-sm outline-none ring-cyan-400/60 transition focus:ring"
              />
              <input
                type="number"
                min={1}
                value={nextPalletTarget}
                onChange={(e) => setNextPalletTarget(e.target.value)}
                placeholder="Nueva paquetería"
                className="rounded-xl border border-white/20 bg-slate-900/70 px-3 py-2 text-sm outline-none ring-cyan-400/60 transition focus:ring"
              />
            </div>

            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => {
                  setShowCapturePalletModal(false);
                  setNextPalletLpn("");
                  setNextPalletTarget("1");
                }}
                className="rounded-xl border border-white/20 bg-white/10 px-3 py-2 text-xs font-medium"
              >
                Cancelar
              </button>
              <button
                type="button"
                onClick={captureNextPallet}
                className="rounded-xl border border-cyan-300/40 bg-cyan-500/20 px-3 py-2 text-xs font-semibold text-cyan-100"
              >
                Iniciar lectura
              </button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}
