"use client";

import { useEffect, useMemo, useRef, useState } from "react";

type Batch = {
  id: number;
  fileName: string;
  totalRows: number;
  validRows: number;
  createdAt: string | Date;
};

type Session = {
  id: number;
  importBatchId: number;
  lpn: string;
  sku: string;
  lastCodigoSap?: string | null;
  lastDescripcion?: string | null;
  packaging: string;
  targetQuantity: number;
  status: string;
  startedAt: string | Date;
  finishedAt: string | Date | null;
};

type ScanEvent = {
  id: number;
  scannedSeries: string;
  matchedSeries: string | null;
  serie1: string | null;
  serie2: string | null;
  codigoSap: string | null;
  descripcion: string | null;
  cantidad: number | null;
  status: "ok" | "duplicate" | "not_found";
  message: string;
  createdAt: string | Date;
};

type Props = {
  initialBatches: Batch[];
};

const badgeClassByStatus: Record<ScanEvent["status"], string> = {
  ok: "bg-emerald-100 text-emerald-700 border-emerald-300",
  duplicate: "bg-amber-100 text-amber-700 border-amber-300",
  not_found: "bg-rose-100 text-rose-700 border-rose-300",
};

const statusLabel: Record<ScanEvent["status"], string> = {
  ok: "Válida",
  duplicate: "Repetida",
  not_found: "No encontrada",
};

const playAlertSound = (kind: "duplicate" | "not_found") => {
  if (typeof window === "undefined") return;

  const context = new window.AudioContext();
  const oscillator = context.createOscillator();
  const gain = context.createGain();

  oscillator.type = "square";
  oscillator.frequency.value = kind === "duplicate" ? 420 : 260;
  gain.gain.value = 0.001;

  oscillator.connect(gain);
  gain.connect(context.destination);

  const now = context.currentTime;
  gain.gain.exponentialRampToValueAtTime(0.28, now + 0.02);
  gain.gain.exponentialRampToValueAtTime(0.001, now + 0.22);

  if (kind === "not_found") {
    oscillator.frequency.setValueAtTime(260, now);
    oscillator.frequency.setValueAtTime(220, now + 0.08);
    oscillator.frequency.setValueAtTime(180, now + 0.14);
  }

  oscillator.start(now);
  oscillator.stop(now + 0.24);

  if (navigator.vibrate) {
    navigator.vibrate(kind === "duplicate" ? [80, 50, 80] : [120, 80, 120]);
  }

  void context.resume();
};

export function SeriesValidatorApp({ initialBatches }: Props) {
  const [batches, setBatches] = useState<Batch[]>(initialBatches);
  const [selectedBatchId, setSelectedBatchId] = useState<number | "">(initialBatches[0]?.id ?? "");

  const [uploading, setUploading] = useState(false);
  const [clearingImports, setClearingImports] = useState(false);
  const [scanLoading, setScanLoading] = useState(false);
  const [startingSession, setStartingSession] = useState(false);
  const [finishingSession, setFinishingSession] = useState(false);

  const [lpn, setLpn] = useState("");
  const [packaging, setPackaging] = useState("1");

  const [activeSession, setActiveSession] = useState<Session | null>(null);
  const [scanValue, setScanValue] = useState("");
  const [events, setEvents] = useState<ScanEvent[]>([]);
  const [palletSessions, setPalletSessions] = useState<Session[]>([]);
  const [alert, setAlert] = useState<{ type: "success" | "warning" | "error"; text: string } | null>(null);
  const [counters, setCounters] = useState<Record<string, number>>({});
  const [duplicatePending, setDuplicatePending] = useState<ScanEvent | null>(null);
  const [exportingSession, setExportingSession] = useState(false);
  const [exportingAll, setExportingAll] = useState(false);
  const [isPastePreviewing, setIsPastePreviewing] = useState(false);

  const fileRef = useRef<HTMLInputElement | null>(null);
  const scanRef = useRef<HTMLInputElement | null>(null);
  const autoTimerRef = useRef<number | null>(null);
  const pastePreviewTimerRef = useRef<number | null>(null);

  const selectedBatch = useMemo(
    () => batches.find((batch) => batch.id === selectedBatchId),
    [batches, selectedBatchId],
  );

  const okReads = counters.ok ?? 0;
  const targetReads = activeSession?.targetQuantity ?? 0;
  const progressPercent = targetReads > 0 ? Math.min(100, Math.round((okReads / targetReads) * 100)) : 0;
  const lastEvent = events[0] ?? null;
  const preferredSessionId = activeSession?.id ?? palletSessions[0]?.id;
  const preferredSessionExportHref = preferredSessionId
    ? `/api/sessions/${preferredSessionId}/export`
    : "#";
  const allPalletsExportHref = selectedBatchId
    ? `/api/exports/pallets?batchId=${selectedBatchId}`
    : "/api/exports/pallets";

  const refreshBatches = async () => {
    const response = await fetch("/api/imports", { cache: "no-store" });
    const data = (await response.json()) as { batches?: Batch[] };
    if (response.ok && data.batches) {
      setBatches(data.batches);
      if (!selectedBatchId && data.batches.length) {
        setSelectedBatchId(data.batches[0].id);
      }
    }
  };

  const refreshSessionDetails = async (sessionId: number) => {
    const response = await fetch(`/api/sessions/${sessionId}`, { cache: "no-store" });
    const data = (await response.json()) as {
      session?: Session;
      events?: ScanEvent[];
      counters?: Array<{ status: string; total: number }>;
    };

    if (!response.ok || !data.session) return;

    setActiveSession(data.session);
    setEvents(data.events ?? []);

    const nextCounters: Record<string, number> = {};
    for (const row of data.counters ?? []) {
      nextCounters[row.status] = Number(row.total);
    }
    setCounters(nextCounters);
  };

  const refreshPalletSessions = async () => {
    const response = await fetch("/api/sessions", { cache: "no-store" });
    const data = (await response.json()) as { sessions?: Session[] };

    if (response.ok && data.sessions) {
      setPalletSessions(data.sessions);
    }
  };

  const openDirectDownload = (url: string) => {
    const link = document.createElement("a");
    link.href = url;
    link.rel = "noopener";
    document.body.appendChild(link);
    link.click();
    link.remove();
  };

  const triggerExcelExport = async (sessionId: number) => {
    setExportingSession(true);
    try {
      openDirectDownload(`/api/sessions/${sessionId}/export`);
      setAlert({ type: "success", text: `Descargando Excel del pallet ${sessionId}...` });
    } catch {
      setAlert({ type: "error", text: "No se pudo iniciar la descarga del Excel." });
    } finally {
      setExportingSession(false);
    }
  };

  const exportPreferredSession = async () => {
    const sessionId = activeSession?.id ?? palletSessions[0]?.id;

    if (!sessionId) {
      setAlert({ type: "error", text: "No hay pallets disponibles para exportar." });
      return;
    }

    await triggerExcelExport(sessionId);
  };

  const exportAllPallets = async () => {
    setExportingAll(true);
    try {
      const batchFilter = selectedBatchId ? `?batchId=${selectedBatchId}` : "";
      openDirectDownload(`/api/exports/pallets${batchFilter}`);
      setAlert({ type: "success", text: "Descargando Excel consolidado de pallets..." });
    } catch {
      setAlert({ type: "error", text: "No se pudo iniciar la descarga consolidada." });
    } finally {
      setExportingAll(false);
    }
  };

  const openSessionForConsultation = async (session: Session) => {
    setActiveSession(session);
    await refreshSessionDetails(session.id);
    setAlert({ type: "success", text: `Consulta cargada para pallet LPN ${session.lpn}.` });
  };

  const finishSession = async (options?: { auto?: boolean }) => {
    if (!activeSession || activeSession.status !== "active") return;

    setFinishingSession(true);
    try {
      const response = await fetch(`/api/sessions/${activeSession.id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "finish" }),
      });

      const data = (await response.json()) as { error?: string; session?: Session };
      if (!response.ok || !data.session) {
        setAlert({ type: "error", text: data.error ?? "No se pudo finalizar la lectura." });
        return;
      }

      setActiveSession(data.session);
      await refreshPalletSessions();

      if (options?.auto) {
        setAlert({
          type: "success",
          text: "Paquetería finalizada automáticamente. Exportando Excel...",
        });
        await triggerExcelExport(data.session.id);
      } else {
        setAlert({
          type: "success",
          text: "Lectura finalizada. Puedes exportar a Excel o continuar con otro pallet.",
        });
      }
    } finally {
      setFinishingSession(false);
    }
  };

  const handleUpload = async (file: File) => {
    setUploading(true);
    setAlert(null);

    try {
      const formData = new FormData();
      formData.append("file", file);

      const response = await fetch("/api/imports", { method: "POST", body: formData });
      const data = (await response.json()) as { error?: string; message?: string; batch?: Batch };

      if (!response.ok) {
        setAlert({ type: "error", text: data.error ?? "No se pudo importar el archivo." });
        return;
      }

      setAlert({ type: "success", text: data.message ?? "Importación completada." });
      await refreshBatches();
      if (data.batch) setSelectedBatchId(data.batch.id);
    } finally {
      setUploading(false);
      if (fileRef.current) fileRef.current.value = "";
    }
  };

  const handleClearImports = async () => {
    if (clearingImports) return;

    setClearingImports(true);
    setAlert(null);

    try {
      const response = await fetch("/api/imports", { method: "DELETE" });
      const data = (await response.json()) as { error?: string; message?: string };

      if (!response.ok) {
        setAlert({ type: "error", text: data.error ?? "No se pudo limpiar la importación." });
        return;
      }

      setBatches([]);
      setSelectedBatchId("");
      setActiveSession(null);
      setEvents([]);
      setPalletSessions([]);
      setCounters({});
      setDuplicatePending(null);
      setScanValue("");
      setLpn("");
      setPackaging("1");

      setAlert({ type: "success", text: data.message ?? "Importaciones limpiadas." });
      await refreshBatches();
      await refreshPalletSessions();
    } finally {
      setClearingImports(false);
    }
  };

  const startNewTask = async () => {
    await handleClearImports();
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const startSession = async () => {
    setStartingSession(true);
    setAlert(null);

    try {
      const response = await fetch("/api/sessions", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          importBatchId: selectedBatchId,
          lpn,
          packaging,
        }),
      });

      const data = (await response.json()) as { error?: string; session?: Session };

      if (!response.ok || !data.session) {
        setAlert({ type: "error", text: data.error ?? "No se pudo iniciar la sesión." });
        return;
      }

      setActiveSession(data.session);
      setEvents([]);
      setCounters({});
      setScanValue("");
      setDuplicatePending(null);
      await refreshPalletSessions();
      setAlert({ type: "success", text: "Sesión iniciada. Puedes escanear o pegar series." });
      setTimeout(() => scanRef.current?.focus(), 50);
    } finally {
      setStartingSession(false);
    }
  };

  const handleScan = async (rawValue?: string) => {
    if (!activeSession || activeSession.status !== "active") return;
    const value = (rawValue ?? scanValue).trim();
    if (!value || scanLoading) return;

    setScanLoading(true);
    setScanValue("");

    try {
      const response = await fetch(`/api/sessions/${activeSession.id}/scan`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ scannedSeries: value }),
      });

      const data = (await response.json()) as {
        error?: string;
        event?: ScanEvent;
        session?: Session;
        lastReads?: ScanEvent[];
        counters?: Record<string, number>;
      };

      if (!response.ok || !data.event) {
        setAlert({ type: "error", text: data.error ?? "Error al procesar el escaneo." });
        return;
      }

      if (data.session) setActiveSession(data.session);
      if (data.lastReads) {
        setEvents(data.lastReads);
      } else {
        setEvents((prev) => [data.event!, ...prev].slice(0, 30));
      }

      const nextCounters = data.counters ?? counters;
      setCounters(nextCounters);

      if (data.event.status === "ok") {
        setAlert({
          type: "success",
          text: `Lectura válida. Serie asociada: ${data.event.matchedSeries ?? "-"}`,
        });

        await refreshPalletSessions();

        const target = data.session?.targetQuantity ?? activeSession.targetQuantity;
        const ok = nextCounters.ok ?? 0;

        if (target > 0 && ok >= target) {
          await finishSession({ auto: true });
        }
      }

      if (data.event.status === "duplicate") {
        playAlertSound("duplicate");
        setDuplicatePending(data.event);
        setAlert({
          type: "warning",
          text: "Serie repetida detectada. Decide si quieres continuar o eliminar esa lectura.",
        });
      }

      if (data.event.status === "not_found") {
        playAlertSound("not_found");
        setAlert({
          type: "error",
          text: "ALERTA: Serie no encontrada. Favor leer la serie correcta.",
        });
      }
    } finally {
      setScanLoading(false);
      setTimeout(() => scanRef.current?.focus(), 60);
    }
  };

  const deleteDuplicateReading = async () => {
    if (!activeSession || !duplicatePending) return;

    const eventId = duplicatePending.id;
    const response = await fetch(`/api/sessions/${activeSession.id}/events/${eventId}`, {
      method: "DELETE",
    });

    const data = (await response.json()) as {
      error?: string;
      counters?: Record<string, number>;
    };

    if (!response.ok) {
      setAlert({ type: "error", text: data.error ?? "No se pudo eliminar la lectura repetida." });
      return;
    }

    setDuplicatePending(null);
    setEvents((prev) => prev.filter((item) => item.id !== eventId));
    if (data.counters) setCounters(data.counters);
    setAlert({ type: "success", text: "Lectura repetida eliminada correctamente." });

    await refreshSessionDetails(activeSession.id);
  };

  const continueDuplicateReading = () => {
    setDuplicatePending(null);
    setAlert({ type: "warning", text: "Se mantuvo el registro de serie repetida para trazabilidad." });
    setTimeout(() => scanRef.current?.focus(), 50);
  };

  const resetForNextPallet = async () => {
    setActiveSession(null);
    setEvents([]);
    setCounters({});
    setScanValue("");
    setDuplicatePending(null);
    setAlert(null);
    setLpn("");
    setPackaging("1");
    await refreshPalletSessions();
  };

  useEffect(() => {
    if (
      !activeSession ||
      activeSession.status !== "active" ||
      !scanValue.trim() ||
      scanLoading ||
      isPastePreviewing
    ) {
      return;
    }

    if (autoTimerRef.current) window.clearTimeout(autoTimerRef.current);

    autoTimerRef.current = window.setTimeout(() => {
      void handleScan(scanValue);
    }, 280);

    return () => {
      if (autoTimerRef.current) window.clearTimeout(autoTimerRef.current);
    };
  }, [scanValue, activeSession, scanLoading, isPastePreviewing]);

  useEffect(() => {
    void refreshPalletSessions();
  }, []);

  useEffect(() => {
    return () => {
      if (pastePreviewTimerRef.current) {
        window.clearTimeout(pastePreviewTimerRef.current);
      }
    };
  }, []);

  return (
    <main className="min-h-screen bg-gradient-to-br from-slate-950 via-slate-900 to-sky-950 p-4 text-slate-100 md:p-6">
      <div className="mx-auto grid w-full max-w-7xl gap-4">
        <header className="rounded-2xl border border-white/10 bg-white/5 p-3 shadow-xl backdrop-blur">
          <p className="text-[10px] uppercase tracking-[0.18em] text-sky-300">Plataforma logística</p>
          <h1 className="mt-1 text-xl font-semibold md:text-2xl">Validación y Asociación de Series</h1>
          <p className="mt-1 max-w-3xl text-xs text-slate-300 md:text-sm">
            Importa planillas masivas y comienza la lectura de series rápidamente.
          </p>
        </header>

        <section className="grid gap-4 md:grid-cols-[1.15fr_1fr]">
          <article className="rounded-3xl border border-white/10 bg-white/5 p-6 shadow-xl backdrop-blur">
            <h2 className="text-lg font-semibold text-white">1) Importación de planilla Excel</h2>
            <p className="mt-2 text-sm text-slate-300">
              Formato: <span className="font-medium text-slate-100">serie 1, serie 2, CODIGOSAP, DESCRIPCION, cantidad</span>
            </p>

            <div className="mt-4 flex flex-wrap items-center gap-3">
              <input
                ref={fileRef}
                type="file"
                accept=".xlsx,.xls"
                className="block rounded-xl border border-white/20 bg-slate-900/70 px-3 py-2 text-sm"
                onChange={(event) => {
                  const file = event.target.files?.[0];
                  if (file) void handleUpload(file);
                }}
              />
              <button
                type="button"
                onClick={() => void handleClearImports()}
                disabled={clearingImports || uploading}
                className="rounded-xl border border-rose-300/40 bg-rose-500/20 px-3 py-2 text-xs font-medium text-rose-100 transition hover:bg-rose-500/30 disabled:cursor-not-allowed disabled:opacity-50"
              >
                {clearingImports ? "Limpiando..." : "Limpiar archivo"}
              </button>
              {uploading && <span className="text-sm text-sky-300">Importando archivo...</span>}
            </div>

            <div className="mt-5 max-h-56 overflow-auto rounded-2xl border border-white/10">
              <table className="w-full text-left text-xs">
                <thead className="bg-white/10 text-slate-200">
                  <tr>
                    <th className="px-3 py-2">Archivo</th>
                    <th className="px-3 py-2">Filas</th>
                    <th className="px-3 py-2">Válidas</th>
                    <th className="px-3 py-2">Fecha</th>
                  </tr>
                </thead>
                <tbody>
                  {batches.map((batch) => (
                    <tr
                      key={batch.id}
                      className={`cursor-pointer border-t border-white/5 ${
                        selectedBatchId === batch.id ? "bg-sky-500/20" : "hover:bg-white/5"
                      }`}
                      onClick={() => setSelectedBatchId(batch.id)}
                    >
                      <td className="px-3 py-2">{batch.fileName}</td>
                      <td className="px-3 py-2">{batch.totalRows.toLocaleString("es-CL")}</td>
                      <td className="px-3 py-2">{batch.validRows.toLocaleString("es-CL")}</td>
                      <td className="px-3 py-2">{new Date(batch.createdAt).toLocaleString("es-CL")}</td>
                    </tr>
                  ))}
                  {!batches.length && (
                    <tr>
                      <td className="px-3 py-4 text-slate-400" colSpan={4}>
                        Aún no hay importaciones.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </article>

          <article className="rounded-3xl border border-white/10 bg-white/5 p-6 shadow-xl backdrop-blur">
            <h2 className="text-lg font-semibold text-white">2) Configuración del pallet</h2>
            <div className="mt-4 grid gap-3">
              <label className="grid gap-1 text-sm">
                <span className="text-slate-300">LPN</span>
                <input
                  value={lpn}
                  onChange={(e) => setLpn(e.target.value)}
                  disabled={Boolean(activeSession && activeSession.status === "active")}
                  className="rounded-xl border border-white/20 bg-slate-900/70 px-3 py-2 outline-none ring-sky-400 transition focus:ring"
                  placeholder="Ej: LPNA00123"
                />
              </label>
              <div className="rounded-xl border border-sky-300/20 bg-sky-500/10 px-3 py-2 text-xs text-sky-100">
                SKU automático: se completa con la primera lectura válida de serie.
              </div>
              <label className="grid gap-1 text-sm">
                <span className="text-slate-300">Paquetería pallet</span>
                <input
                  type="number"
                  min={1}
                  value={packaging}
                  onChange={(e) => setPackaging(e.target.value)}
                  disabled={Boolean(activeSession && activeSession.status === "active")}
                  className="rounded-xl border border-white/20 bg-slate-900/70 px-3 py-2 outline-none ring-sky-400 transition focus:ring"
                  placeholder="Ej: 1200"
                />
              </label>
            </div>

            <button
              type="button"
              disabled={
                startingSession ||
                !selectedBatch ||
                Boolean(activeSession && activeSession.status === "active")
              }
              onClick={() => void startSession()}
              className="mt-5 w-full rounded-xl bg-sky-500 px-4 py-2.5 font-medium text-white transition hover:bg-sky-400 disabled:cursor-not-allowed disabled:opacity-50"
            >
              {startingSession ? "Iniciando sesión..." : "Iniciar lectura de pallet"}
            </button>
          </article>
        </section>

        {alert && (
          <div
            className={`rounded-2xl border px-4 py-3 text-sm ${
              alert.type === "success"
                ? "border-emerald-300/50 bg-emerald-500/15 text-emerald-100"
                : alert.type === "warning"
                  ? "border-amber-300/50 bg-amber-500/15 text-amber-100"
                  : "border-rose-300/50 bg-rose-500/15 text-rose-100"
            }`}
          >
            {alert.text}
          </div>
        )}

        <section className="grid gap-6 lg:grid-cols-[1fr_1.3fr]">
          <article className="rounded-3xl border border-white/10 bg-white/5 p-6 shadow-xl backdrop-blur">
            <h2 className="text-lg font-semibold text-white">3) Lectura de series</h2>
            <p className="mt-2 text-sm text-slate-300">
              Pega o escanea una serie. La validación es automática, sin presionar Enter ni botón.
            </p>

            <div className="mt-4 grid gap-3">
              <input
                ref={scanRef}
                value={scanValue}
                onChange={(e) => setScanValue(e.target.value)}
                onPaste={(event) => {
                  event.preventDefault();
                  const pasted = event.clipboardData.getData("text").trim();
                  setScanValue(pasted);

                  if (!pasted) return;

                  setIsPastePreviewing(true);
                  if (pastePreviewTimerRef.current) {
                    window.clearTimeout(pastePreviewTimerRef.current);
                  }

                  pastePreviewTimerRef.current = window.setTimeout(() => {
                    setIsPastePreviewing(false);
                    void handleScan(pasted);
                  }, 450);
                }}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    e.preventDefault();
                    void handleScan();
                  }
                }}
                disabled={!activeSession || activeSession.status !== "active"}
                className="rounded-xl border border-white/20 bg-slate-900/80 px-3 py-3 text-lg outline-none ring-sky-400 transition focus:ring disabled:cursor-not-allowed disabled:opacity-50"
                placeholder="Escanea o pega una serie..."
              />
              <button
                type="button"
                onClick={() => void finishSession()}
                disabled={!activeSession || activeSession.status !== "active" || finishingSession}
                className="rounded-xl bg-emerald-500 px-4 py-2.5 font-medium text-white transition hover:bg-emerald-400 disabled:cursor-not-allowed disabled:opacity-50"
              >
                {finishingSession ? "Finalizando..." : "Finalizar lectura"}
              </button>
            </div>

            <div className="mt-5 rounded-2xl border border-sky-300/30 bg-sky-500/10 p-3 text-sm">
              <div className="mb-2 flex items-center justify-between">
                <p>
                  Progreso paquetería: <strong>{okReads}</strong> / <strong>{targetReads || "-"}</strong>
                </p>
                <p className="text-sky-100">{progressPercent}%</p>
              </div>
              <div className="h-2 w-full rounded-full bg-slate-800/80">
                <div
                  className="h-2 rounded-full bg-sky-400 transition-all"
                  style={{ width: `${progressPercent}%` }}
                />
              </div>
              {activeSession?.status === "finished" && (
                <p className="mt-2 font-medium text-emerald-200">Finalizada ✅</p>
              )}
            </div>

            <div className="mt-5 grid grid-cols-3 gap-3 text-center text-sm">
              <div className="rounded-2xl border border-emerald-400/40 bg-emerald-500/10 p-3">
                <p className="text-2xl font-semibold">{counters.ok ?? 0}</p>
                <p className="text-emerald-100">Válidas</p>
              </div>
              <div className="rounded-2xl border border-amber-400/40 bg-amber-500/10 p-3">
                <p className="text-2xl font-semibold">{counters.duplicate ?? 0}</p>
                <p className="text-amber-100">Repetidas</p>
              </div>
              <div className="rounded-2xl border border-rose-400/40 bg-rose-500/10 p-3">
                <p className="text-2xl font-semibold">{counters.not_found ?? 0}</p>
                <p className="text-rose-100">No encontradas</p>
              </div>
            </div>

            <div className="mt-6 flex flex-wrap gap-3">
              <a
                href={preferredSessionExportHref}
                className={`rounded-xl border border-sky-300/40 bg-sky-500/20 px-4 py-2 text-sm font-medium text-sky-100 transition hover:bg-sky-500/30 ${
                  !preferredSessionId ? "pointer-events-none opacity-50" : ""
                }`}
              >
                Exportar a Excel
              </a>

              <a
                href={allPalletsExportHref}
                className={`rounded-xl border border-emerald-300/40 bg-emerald-500/20 px-4 py-2 text-sm font-medium text-emerald-100 transition hover:bg-emerald-500/30 ${
                  !palletSessions.length ? "pointer-events-none opacity-50" : ""
                }`}
              >
                Exportar todos los pallets
              </a>

              <button
                type="button"
                onClick={() => void resetForNextPallet()}
                className="rounded-xl border border-white/20 bg-white/10 px-4 py-2 text-sm font-medium transition hover:bg-white/20"
              >
                Continuar con otro pallet
              </button>
            </div>
          </article>

          <article className="rounded-3xl border border-white/10 bg-white/5 p-6 shadow-xl backdrop-blur">
            <h2 className="text-lg font-semibold text-white">Última lectura e historial</h2>

            <div className="mt-4 rounded-2xl border border-white/10 bg-slate-900/50 p-4">
              {lastEvent ? (
                <div className="space-y-2 text-sm">
                  <div className="flex items-center justify-between">
                    <span className={`rounded-full border px-2.5 py-0.5 text-xs ${badgeClassByStatus[lastEvent.status]}`}>
                      {statusLabel[lastEvent.status]}
                    </span>
                    <span className="text-slate-400">{new Date(lastEvent.createdAt).toLocaleTimeString("es-CL")}</span>
                  </div>
                  <p className="flex flex-wrap items-center gap-x-4 gap-y-1">
                    <span>
                      <span className="text-slate-400">Series 1:</span>{" "}
                      <strong
                        className={
                          lastEvent.scannedSeries === lastEvent.serie1
                            ? "rounded bg-sky-500/25 px-2 py-0.5 text-sky-100"
                            : ""
                        }
                      >
                        {lastEvent.serie1 ?? "-"}
                      </strong>
                    </span>
                    <span>
                      <span className="text-slate-400">Series 2:</span>{" "}
                      <strong
                        className={
                          lastEvent.scannedSeries === lastEvent.serie2
                            ? "rounded bg-sky-500/25 px-2 py-0.5 text-sky-100"
                            : ""
                        }
                      >
                        {lastEvent.serie2 ?? "-"}
                      </strong>
                    </span>
                  </p>
                  <p>
                    <span className="text-slate-400">SKU/Código:</span> {lastEvent.codigoSap ?? "-"}
                  </p>
                  <p className="text-slate-300">{lastEvent.message}</p>
                </div>
              ) : (
                <p className="text-sm text-slate-400">Sin lecturas todavía.</p>
              )}
            </div>

            <div className="mt-4 max-h-72 overflow-auto rounded-2xl border border-white/10">
              <table className="w-full text-left text-xs">
                <thead className="bg-white/10 text-slate-200">
                  <tr>
                    <th className="px-3 py-2">Estado</th>
                    <th className="px-3 py-2">Series 1</th>
                    <th className="px-3 py-2">Series 2</th>
                    <th className="px-3 py-2">Hora</th>
                  </tr>
                </thead>
                <tbody>
                  {events.map((event) => (
                    <tr key={event.id} className="border-t border-white/5">
                      <td className="px-3 py-2">
                        <span className={`rounded-full border px-2 py-0.5 ${badgeClassByStatus[event.status]}`}>
                          {statusLabel[event.status]}
                        </span>
                      </td>
                      <td className="px-3 py-2">
                        <span
                          className={
                            event.scannedSeries === event.serie1
                              ? "rounded bg-sky-500/25 px-2 py-0.5 text-sky-100"
                              : ""
                          }
                        >
                          {event.serie1 ?? "-"}
                        </span>
                      </td>
                      <td className="px-3 py-2">
                        <span
                          className={
                            event.scannedSeries === event.serie2
                              ? "rounded bg-sky-500/25 px-2 py-0.5 text-sky-100"
                              : ""
                          }
                        >
                          {event.serie2 ?? "-"}
                        </span>
                      </td>
                      <td className="px-3 py-2">{new Date(event.createdAt).toLocaleTimeString("es-CL")}</td>
                    </tr>
                  ))}
                  {!events.length && (
                    <tr>
                      <td className="px-3 py-4 text-slate-400" colSpan={4}>
                        El historial de lecturas aparecerá aquí.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </article>
        </section>

        <section className="rounded-3xl border border-white/10 bg-white/5 p-4 shadow-xl backdrop-blur">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <h3 className="text-base font-semibold text-white">Consulta de pallets registrados</h3>
            <p className="text-xs text-slate-300">Total visibles: {palletSessions.length}</p>
          </div>

          <div className="mt-3 max-h-56 overflow-auto rounded-2xl border border-white/10">
            <table className="w-full text-left text-xs">
              <thead className="bg-white/10 text-slate-200">
                <tr>
                  <th className="px-3 py-2">LPN</th>
                  <th className="px-3 py-2">Paquetería</th>
                  <th className="px-3 py-2">SKU</th>
                  <th className="px-3 py-2">Descripción</th>
                  <th className="px-3 py-2">Estado</th>
                  <th className="px-3 py-2">Inicio</th>
                </tr>
              </thead>
              <tbody>
                {palletSessions.map((session) => (
                  <tr
                    key={session.id}
                    className="cursor-pointer border-t border-white/5 hover:bg-white/5"
                    onClick={() => void openSessionForConsultation(session)}
                  >
                    <td className="px-3 py-2">{session.lpn}</td>
                    <td className="px-3 py-2">{session.packaging}</td>
                    <td className="px-3 py-2">{session.lastCodigoSap ?? session.sku}</td>
                    <td className="px-3 py-2">{session.lastDescripcion ?? "Sin lectura válida"}</td>
                    <td className="px-3 py-2">
                      <span
                        className={`rounded-full px-2 py-0.5 ${
                          session.status === "finished"
                            ? "bg-emerald-500/20 text-emerald-100"
                            : "bg-amber-500/20 text-amber-100"
                        }`}
                      >
                        {session.status === "finished" ? "Finalizado" : "Activo"}
                      </span>
                    </td>
                    <td className="px-3 py-2">{new Date(session.startedAt).toLocaleString("es-CL")}</td>
                  </tr>
                ))}
                {!palletSessions.length && (
                  <tr>
                    <td className="px-3 py-4 text-slate-400" colSpan={6}>
                      Aún no hay pallets registrados.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </section>
      </div>

      <button
        type="button"
        onClick={() => void startNewTask()}
        disabled={clearingImports}
        className="fixed right-4 top-4 z-40 rounded-full border border-rose-300/40 bg-rose-500/90 px-4 py-2.5 text-xs font-semibold text-white shadow-2xl transition hover:bg-rose-500 md:right-6 md:top-6 md:text-sm disabled:cursor-not-allowed disabled:opacity-60"
      >
        {clearingImports ? "Reiniciando..." : "Iniciar nueva tarea"}
      </button>

      {duplicatePending && (
        <div className="fixed inset-0 z-50 grid place-items-center bg-slate-950/70 p-4">
          <div className="w-full max-w-md rounded-2xl border border-amber-300/40 bg-slate-900 p-5 shadow-2xl">
            <h3 className="text-lg font-semibold text-amber-200">Serie repetida detectada</h3>
            <p className="mt-2 text-sm text-slate-200">
              Serie: <strong>{duplicatePending.scannedSeries}</strong>
            </p>
            <p className="mt-1 text-sm text-slate-300">
              Esta lectura ya existe en la sesión. ¿Deseas continuar o eliminar esta lectura repetida?
            </p>

            <div className="mt-5 flex justify-end gap-3">
              <button
                type="button"
                onClick={deleteDuplicateReading}
                className="rounded-xl border border-rose-300/40 bg-rose-500/20 px-4 py-2 text-sm font-medium text-rose-100 hover:bg-rose-500/30"
              >
                Eliminar lectura
              </button>
              <button
                type="button"
                onClick={continueDuplicateReading}
                className="rounded-xl border border-amber-300/40 bg-amber-500/20 px-4 py-2 text-sm font-medium text-amber-100 hover:bg-amber-500/30"
              >
                Continuar
              </button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}
