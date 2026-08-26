import { useRef, useState } from "react";
import type { ChangeEvent, DragEvent } from "react";
import {
  AlertCircle,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  Loader2,
  ShieldCheck,
  UploadCloud,
  X,
} from "lucide-react";

type Result = {
  total_rows: number;
  correct_rows: number;
  discarded_rows: number;
  output_filename: string;
  checks: Record<string, number>;
};

const API_URL = "https://upcensus-detector.vercel.app/";

const ACCEPTED = ".xlsx,.xls,.csv,.tsv";

export default function App() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [file, setFile] = useState<File | null>(null);
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<Result | null>(null);
  const [error, setError] = useState("");

  const selectFile = (selected: File | null) => {
    setError("");
    setResult(null);

    if (!selected) {
      setFile(null);
      return;
    }

    const allowed = [".xlsx", ".xls", ".csv", ".tsv"];
    const lower = selected.name.toLowerCase();

    if (!allowed.some((ext) => lower.endsWith(ext))) {
      setError("Please upload an Excel, CSV, or TSV file.");
      setFile(null);
      return;
    }

    setFile(selected);
  };

  const onInputChange = (event: ChangeEvent<HTMLInputElement>) => {
    selectFile(event.target.files?.[0] ?? null);
  };

  const onDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragging(false);
    selectFile(event.dataTransfer.files?.[0] ?? null);
  };

  const validate = async () => {
    if (!file) return;

    setLoading(true);
    setError("");
    setResult(null);

    try {
      const form = new FormData();
      form.append("file", file);

      const response = await fetch(`${API_URL}/validate`, {
        method: "POST",
        body: form,
      });

      if (!response.ok) {
        const data: unknown = await response.json().catch(() => null);
        const message =
          typeof data === "object" &&
          data !== null &&
          "detail" in data &&
          typeof data.detail === "string"
            ? data.detail
            : "Validation failed.";
        throw new Error(message);
      }

      const data = (await response.json()) as Result;
      setResult(data);
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Could not connect to the Python backend.",
      );
    } finally {
      setLoading(false);
    }
  };

  const download = () => {
    if (!result) return;
    window.location.href = `${API_URL}/download/${encodeURIComponent(result.output_filename)}`;
  };

  const reset = () => {
    setFile(null);
    setResult(null);
    setError("");
    if (inputRef.current) inputRef.current.value = "";
  };

  return (
    <main className="min-h-screen bg-slate-950 text-white">
      <div className="mx-auto max-w-6xl px-5 py-10 sm:px-8">
        <header className="mb-8">
          <div className="mb-4 flex items-center gap-3">
            <div className="rounded-2xl bg-blue-600 p-3 shadow-lg shadow-blue-600/20">
              <ShieldCheck size={25} />
            </div>
            <div>
              <h1 className="text-2xl font-bold tracking-tight sm:text-3xl">
                Survey Spreadsheet Validator
              </h1>
              <p className="text-sm text-slate-400">
                Upload your survey data and separate valid and problematic records.
              </p>
            </div>
          </div>
        </header>

        <section className="grid gap-6 lg:grid-cols-[1.4fr_0.6fr]">
          <div className="rounded-3xl border border-slate-800 bg-slate-900/80 p-6 shadow-2xl">
            <div
              onDragOver={(event) => {
                event.preventDefault();
                setDragging(true);
              }}
              onDragLeave={() => setDragging(false)}
              onDrop={onDrop}
              onClick={() => inputRef.current?.click()}
              className={`cursor-pointer rounded-2xl border-2 border-dashed p-10 text-center transition ${
                dragging
                  ? "border-blue-400 bg-blue-500/10"
                  : "border-slate-700 hover:border-slate-500 hover:bg-slate-800/60"
              }`}
            >
              <input
                ref={inputRef}
                type="file"
                accept={ACCEPTED}
                className="hidden"
                onChange={onInputChange}
              />
              <UploadCloud className="mx-auto mb-4 text-blue-400" size={42} />
              <h2 className="text-lg font-semibold">
                Drop your spreadsheet here
              </h2>
              <p className="mt-2 text-sm text-slate-400">
                XLSX, XLS, CSV or TSV
              </p>

              {file && (
                <div className="mx-auto mt-6 flex max-w-md items-center gap-3 rounded-xl border border-slate-700 bg-slate-950/70 p-3 text-left">
                  <FileSpreadsheet className="shrink-0 text-emerald-400" />
                  <div className="min-w-0 flex-1">
                    <p className="truncate text-sm font-medium">{file.name}</p>
                    <p className="text-xs text-slate-500">
                      {(file.size / 1024 / 1024).toFixed(2)} MB
                    </p>
                  </div>
                  <button
                    type="button"
                    title="Remove file"
                    onClick={(event) => {
                      event.stopPropagation();
                      reset();
                    }}
                    className="rounded-lg p-2 text-slate-400 hover:bg-slate-800 hover:text-white"
                  >
                    <X size={18} />
                  </button>
                </div>
              )}
            </div>

            {error && (
              <div className="mt-5 flex items-start gap-3 rounded-xl border border-red-900/60 bg-red-950/40 p-4 text-sm text-red-300">
                <AlertCircle className="mt-0.5 shrink-0" size={18} />
                <span>{error}</span>
              </div>
            )}

            <div className="mt-6 flex flex-col gap-3 sm:flex-row">
              <button
                type="button"
                disabled={!file || loading}
                onClick={validate}
                className="flex flex-1 items-center justify-center gap-2 rounded-xl bg-blue-600 px-5 py-3 font-semibold transition hover:bg-blue-500 disabled:cursor-not-allowed disabled:opacity-40"
              >
                {loading ? (
                  <>
                    <Loader2 className="animate-spin" size={19} />
                    Validating...
                  </>
                ) : (
                  <>
                    <ShieldCheck size={19} />
                    Validate Data
                  </>
                )}
              </button>

              {file && (
                <button
                  type="button"
                  onClick={reset}
                  disabled={loading}
                  className="rounded-xl border border-slate-700 px-5 py-3 font-semibold text-slate-300 hover:bg-slate-800 disabled:opacity-40"
                >
                  Clear
                </button>
              )}
            </div>
          </div>

          <aside className="rounded-3xl border border-slate-800 bg-slate-900/80 p-6">
            <h2 className="mb-5 font-semibold">Checks included</h2>
            <div className="space-y-3 text-sm text-slate-300">
              {[
                "Indian mobile number format",
                "Duplicate & suspicious mobile numbers",
                "Duplicate outlet names",
                "Duplicate image references",
                "Screenshot-like image names",
                "Latitude / longitude presence",
                "Latitude / longitude ranges",
              ].map((item) => (
                <div key={item} className="flex gap-3">
                  <CheckCircle2 className="mt-0.5 shrink-0 text-emerald-400" size={17} />
                  <span>{item}</span>
                </div>
              ))}
            </div>
          </aside>
        </section>

        {result && (
          <section className="mt-6 rounded-3xl border border-slate-800 bg-slate-900/80 p-6">
            <div className="mb-5 flex flex-col justify-between gap-4 sm:flex-row sm:items-center">
              <div>
                <h2 className="text-xl font-bold">Validation complete</h2>
                <p className="mt-1 text-sm text-slate-400">
                  Two sheets have been created in the output workbook.
                </p>
              </div>
              <button
                type="button"
                onClick={download}
                className="flex items-center justify-center gap-2 rounded-xl bg-emerald-600 px-5 py-3 font-semibold hover:bg-emerald-500"
              >
                <Download size={19} />
                Download Excel
              </button>
            </div>

            <div className="grid gap-3 sm:grid-cols-3">
              <Stat label="Total Records" value={result.total_rows} />
              <Stat
                label="Correct Data"
                value={result.correct_rows}
                good
              />
              <Stat
                label="Discard Data"
                value={result.discarded_rows}
                bad
              />
            </div>

            <div className="mt-5 grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
              {Object.entries(result.checks).map(([key, value]) => (
                <div
                  key={key}
                  className="rounded-xl border border-slate-800 bg-slate-950/50 p-4"
                >
                  <p className="text-xs uppercase tracking-wide text-slate-500">
                    {key}
                  </p>
                  <p className="mt-1 text-xl font-bold">{value}</p>
                </div>
              ))}
            </div>
          </section>
        )}
      </div>
    </main>
  );
}

function Stat({
  label,
  value,
  good = false,
  bad = false,
}: {
  label: string;
  value: number;
  good?: boolean;
  bad?: boolean;
}) {
  return (
    <div className="rounded-2xl border border-slate-800 bg-slate-950/50 p-5">
      <p className="text-sm text-slate-400">{label}</p>
      <p
        className={`mt-1 text-3xl font-bold ${
          good ? "text-emerald-400" : bad ? "text-red-400" : "text-white"
        }`}
      >
        {value.toLocaleString()}
      </p>
    </div>
  );
}
