"use client";

import React, { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  Search,
  Download,
  ClipboardList,
  CheckCircle2,
  Clock,
  Wrench,
  AlertTriangle,
  FileSearch,
  Info,
  X,
} from "lucide-react";

// Data sumber (hasil konversi PROGRES_per_24_Agustus_2026.xlsx)
import RAW_DATA from "./data/progres-data.json";

export interface ProgresRecord {
  no: number | string;
  area_intervensi: string;
  opd: string;
  data_dokumen_yg_diperlukan: string;
  komitmen_berita_acara_13_agustus_2026: string;
  progress_per_24_agustus_2026: string;
}

// Kategori status hasil pengelompokan teks bebas pada kolom
// "PROGRESS PER 24 AGUSTUS 2026"
type StatusKey =
  | "DITERIMA"
  | "DISIAPKAN"
  | "PERBAIKAN"
  | "REVIU"
  | "BELUM"
  | "TTD"
  | "SURAT_BELUM"
  | "SEBAGIAN"
  | "LAINNYA";

type StatusFilter = "ALL" | StatusKey;

// Mengelompokkan teks progres bebas menjadi kategori status yang konsisten.
// Urutan pengecekan penting: kategori yang lebih spesifik dicek lebih dulu.
function classifyStatus(progress: string): StatusKey {
  const t = (progress || "").toUpperCase();
  if (t.includes("SUDAH DITERIMA")) return "DITERIMA";
  if (t.includes("SEDANG DISIAPKAN")) return "DISIAPKAN";
  if (t.includes("MENUNGGU TANDA TANGAN")) return "TTD";
  if (t.includes("SURAT PERNYATAAN BELUM DI TERIMA")) return "SURAT_BELUM";
  if (t.includes("SEBAGIAN DATA") || t.includes("SEBAGIAN DOKUMEN")) return "SEBAGIAN";
  if (t.includes("PERBAIKAN")) return "PERBAIKAN";
  if (t.includes("REVIU")) return "REVIU";
  if (t.includes("BELUM ADA")) return "BELUM";
  return "LAINNYA";
}

const STATUS_CONFIG: Record<StatusKey, { label: string; badge: string; dot: string }> = {
  DITERIMA: {
    label: "Sudah Diterima",
    badge: "border-emerald-200 bg-emerald-50 text-emerald-700",
    dot: "bg-emerald-500",
  },
  DISIAPKAN: {
    label: "Sedang Disiapkan",
    badge: "border-sky-200 bg-sky-50 text-sky-700",
    dot: "bg-sky-500",
  },
  PERBAIKAN: {
    label: "Dalam Perbaikan",
    badge: "border-amber-200 bg-amber-50 text-amber-700",
    dot: "bg-amber-500",
  },
  REVIU: {
    label: "Proses Reviu",
    badge: "border-violet-200 bg-violet-50 text-violet-700",
    dot: "bg-violet-500",
  },
  BELUM: {
    label: "Belum Ada Progres",
    badge: "border-rose-200 bg-rose-50 text-rose-700",
    dot: "bg-rose-500",
  },
  TTD: {
    label: "Menunggu Tanda Tangan Pejabat",
    badge: "border-fuchsia-200 bg-fuchsia-50 text-fuchsia-700",
    dot: "bg-fuchsia-500",
  },
  SURAT_BELUM: {
    label: "Surat Pernyataan Belum Diterima",
    badge: "border-orange-200 bg-orange-50 text-orange-700",
    dot: "bg-orange-500",
  },
  SEBAGIAN: {
    label: "Sebagian Sudah Diupload",
    badge: "border-teal-200 bg-teal-50 text-teal-700",
    dot: "bg-teal-500",
  },
  LAINNYA: {
    label: "Link Pengaduan Sudah Ada",
    badge: "border-slate-200 bg-slate-50 text-slate-600",
    dot: "bg-slate-400",
  },
};

const FILTER_TABS: { key: StatusFilter; label: string }[] = [
  { key: "ALL", label: "Semua Data" },
  { key: "DITERIMA", label: "Sudah Diterima" },
  { key: "DISIAPKAN", label: "Sedang Disiapkan" },
  { key: "PERBAIKAN", label: "Dalam Perbaikan" },
  { key: "REVIU", label: "Proses Reviu" },
  { key: "BELUM", label: "Belum Ada Progres" },
  { key: "TTD", label: "Menunggu Tanda Tangan Pejabat" },
  { key: "SURAT_BELUM", label: "Surat Pernyataan Belum Diterima" },
  { key: "SEBAGIAN", label: "Sebagian Sudah Diupload" },
  { key: "LAINNYA", label: "Link Pengaduan Sudah Ada" },
];

// Gradient pill per kategori, dipakai untuk state aktif pada filter & kartu statistik
const PILL_GRADIENT: Record<StatusFilter, string> = {
  ALL: "from-indigo-500 to-blue-500",
  DITERIMA: "from-emerald-500 to-teal-500 shadow-emerald-500/25",
  DISIAPKAN: "from-sky-500 to-cyan-500 shadow-sky-500/25",
  PERBAIKAN: "from-amber-500 to-orange-500 shadow-amber-500/25",
  REVIU: "from-violet-500 to-purple-500 shadow-violet-500/25",
  BELUM: "from-rose-500 to-red-500 shadow-rose-500/25",
  TTD: "from-fuchsia-500 to-pink-500 shadow-fuchsia-500/25",
  SURAT_BELUM: "from-orange-500 to-amber-600 shadow-orange-500/25",
  SEBAGIAN: "from-teal-500 to-emerald-600 shadow-teal-500/25",
  LAINNYA: "from-slate-400 to-slate-500 shadow-slate-400/25",
};

export default function ProgresDokumenPage() {
  const rawJson = RAW_DATA as {
    judul?: string;
    per_tanggal?: string;
    jumlah_item?: number;
    data?: ProgresRecord[];
  };

  const data: ProgresRecord[] = Array.isArray(rawJson?.data) ? rawJson.data : [];

  // Indikator scroll horizontal buatan
  const scrollRef = useRef<HTMLDivElement>(null);
  const [scrollBar, setScrollBar] = useState({ widthPct: 100, leftPct: 0 });

  const handleTableScroll = () => {
    const el = scrollRef.current;
    if (!el) return;
    const { scrollLeft, scrollWidth, clientWidth } = el;
    if (scrollWidth <= clientWidth) {
      setScrollBar({ widthPct: 100, leftPct: 0 });
      return;
    }
    const widthPct = Math.max((clientWidth / scrollWidth) * 100, 8);
    const leftPct = (scrollLeft / (scrollWidth - clientWidth)) * (100 - widthPct);
    setScrollBar({ widthPct, leftPct });
  };

  const [searchTerm, setSearchTerm] = useState<string>("");
  const [statusFilter, setStatusFilter] = useState<StatusFilter>("ALL");

  // Hitung jumlah data per kategori status (untuk badge di setiap pill & kartu statistik)
  const statusCounts = useMemo(() => {
    const counts: Record<StatusKey, number> = {
      DITERIMA: 0,
      DISIAPKAN: 0,
      PERBAIKAN: 0,
      REVIU: 0,
      BELUM: 0,
      TTD: 0,
      SURAT_BELUM: 0,
      SEBAGIAN: 0,
      LAINNYA: 0,
    };
    data.forEach((item) => {
      counts[classifyStatus(item.progress_per_24_agustus_2026)] += 1;
    });
    return counts;
  }, [data]);

  const filteredData = useMemo(() => {
    const q = searchTerm.trim().toLowerCase();
    return data.filter((item) => {
      const matchesSearch =
        q === ""
          ? true
          : item.area_intervensi?.toLowerCase().includes(q) ||
            item.opd?.toLowerCase().includes(q) ||
            item.data_dokumen_yg_diperlukan?.toLowerCase().includes(q) ||
            item.progress_per_24_agustus_2026?.toLowerCase().includes(q);

      const matchesStatus =
        statusFilter === "ALL" ? true : classifyStatus(item.progress_per_24_agustus_2026) === statusFilter;

      return matchesSearch && matchesStatus;
    });
  }, [data, searchTerm, statusFilter]);

  const isFiltering = searchTerm.trim().length > 0 || statusFilter !== "ALL";

  const exportToExcel = () => {
    if (filteredData.length === 0) return;

    const exportRows = filteredData.map((d) => ({
      NO: d.no,
      "AREA INTERVENSI": d.area_intervensi,
      OPD: d.opd,
      "DATA / DOKUMEN YG DIPERLUKAN": d.data_dokumen_yg_diperlukan,
      "KOMITMEN BERITA ACARA 13 AGUSTUS 2026": d.komitmen_berita_acara_13_agustus_2026,
      "PROGRESS PER 24 AGUSTUS 2026": d.progress_per_24_agustus_2026,
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    worksheet["!cols"] = [
      { wch: 5 },
      { wch: 28 },
      { wch: 22 },
      { wch: 45 },
      { wch: 28 },
      { wch: 32 },
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Progress Dokumen");
    XLSX.writeFile(workbook, "Progress_Dokumen_MCSP-RBS_KPK_Deiyai.xlsx");
  };

  const resetFilter = () => {
    setSearchTerm("");
    setStatusFilter("ALL");
  };

  return (
    <main className="min-h-screen bg-[radial-gradient(ellipse_at_top,_#eef2ff_0%,_#f8fafc_45%,_#f8fafc_100%)] px-4 py-8 md:px-8">
      <div className="mx-auto max-w-[1500px] space-y-5">
        {/* Header Section — Kop Surat + Hero judul */}
        <div className="overflow-hidden rounded-lg border border-slate-200/70 bg-white shadow-[0_1px_2px_rgba(15,23,42,0.04),0_18px_40px_-20px_rgba(79,70,229,0.25)]">
          {/* Kop surat */}
          <div className="flex flex-col items-center gap-3 px-6 pt-6 text-center sm:flex-row sm:items-center sm:gap-4 sm:px-8 sm:text-left">
            <img
              src="/logo-deiyai.png"
              alt="Logo Kabupaten Deiyai"
              className="h-14 w-14 shrink-0 object-contain sm:h-16 sm:w-16"
            />
            <div>
              <h2 className="text-sm font-bold uppercase tracking-wide text-slate-800 sm:text-base">
                Pemerintah Kabupaten Deiyai
              </h2>
              <p className="text-xs font-medium uppercase tracking-wide text-indigo-600 sm:text-sm">
                Monitoring Center for Prevention (MCSP) - RBS KPK
              </p>
              <p className="mt-0.5 text-[11px] text-slate-400 sm:text-xs">
                Kompleks Perkantoran Pemda Deiyai - Papua Tengah
              </p>
            </div>
          </div>

          {/* Hero judul — gradient banner */}
          <div className="relative mt-6 overflow-hidden bg-gradient-to-r from-indigo-600 via-blue-600 to-sky-500 px-6 py-7 sm:px-10 sm:py-8">
            <div className="pointer-events-none absolute -right-10 -top-10 h-40 w-40 rounded-full bg-white/10 blur-2xl" />
            <div className="pointer-events-none absolute -bottom-16 left-1/3 h-48 w-48 rounded-full bg-white/10 blur-3xl" />

            <div className="relative flex flex-col items-center gap-3 text-center">
              <div className="flex h-12 w-12 items-center justify-center rounded-sm bg-white/15 text-white ring-1 ring-white/30 backdrop-blur">
                <ClipboardList className="h-6 w-6" />
              </div>
              <div>
                <h1 className="text-lg font-extrabold leading-snug tracking-tight text-white sm:text-xl md:text-2xl">
                  Progress Pemenuhan Dokumen MCSP-RBS KPK
                  <br className="hidden sm:block" /> Kabupaten Deiyai Tahun 2026
                </h1>
                {/* <p className="mt-1 text-xs font-medium text-indigo-100 sm:text-sm">
                  {rawJson.per_tanggal ? `Per ${rawJson.per_tanggal}` : "Per 24 Agustus 2026"}
                </p> */}
              </div>
            </div>
          </div>
        </div>

        {/* Ringkasan Statistik (klik untuk memfilter tabel) */}
        <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-6">
          <StatCard
            icon={<ClipboardList className="h-5 w-5" />}
            gradient="from-indigo-500 to-blue-500"
            label="Total Item"
            value={data.length}
            active={statusFilter === "ALL"}
            onClick={() => setStatusFilter("ALL")}
          />
          <StatCard
            icon={<CheckCircle2 className="h-5 w-5" />}
            gradient="from-emerald-500 to-teal-500"
            label="Sudah Diterima"
            value={statusCounts.DITERIMA}
            active={statusFilter === "DITERIMA"}
            onClick={() => setStatusFilter("DITERIMA")}
          />
          <StatCard
            icon={<Clock className="h-5 w-5" />}
            gradient="from-sky-500 to-cyan-500"
            label="Sedang Disiapkan"
            value={statusCounts.DISIAPKAN}
            active={statusFilter === "DISIAPKAN"}
            onClick={() => setStatusFilter("DISIAPKAN")}
          />
          <StatCard
            icon={<Wrench className="h-5 w-5" />}
            gradient="from-amber-500 to-orange-500"
            label="Dalam Perbaikan"
            value={statusCounts.PERBAIKAN}
            active={statusFilter === "PERBAIKAN"}
            onClick={() => setStatusFilter("PERBAIKAN")}
          />
          <StatCard
            icon={<FileSearch className="h-5 w-5" />}
            gradient="from-violet-500 to-purple-500"
            label="Proses Reviu"
            value={statusCounts.REVIU}
            active={statusFilter === "REVIU"}
            onClick={() => setStatusFilter("REVIU")}
          />
          <StatCard
            icon={<AlertTriangle className="h-5 w-5" />}
            gradient="from-rose-500 to-red-500"
            label="Belum Ada Progres"
            value={statusCounts.BELUM}
            active={statusFilter === "BELUM"}
            onClick={() => setStatusFilter("BELUM")}
          />
        </div>

        {/* Input Pencarian & Filter */}
        <div className="space-y-4 rounded-sm border border-slate-200/70 bg-white p-5 shadow-[0_1px_2px_rgba(15,23,42,0.04),0_14px_32px_-18px_rgba(15,23,42,0.14)] sm:p-6">
          <div className="flex flex-col items-stretch justify-between gap-3 md:flex-row md:items-center">
            <div className="relative flex-1">
              <Search className="pointer-events-none absolute left-4 top-1/2 h-4.5 w-4.5 -translate-y-1/2 text-indigo-400" />
              <input
                type="text"
                placeholder="Cari Area Intervensi, OPD, atau Dokumen..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full rounded-sm border border-slate-200 bg-slate-50 py-3 pl-11 pr-10 text-sm font-medium text-slate-800 placeholder:font-normal placeholder:text-slate-400 transition focus:border-indigo-400 focus:bg-white focus:outline-none focus:ring-4 focus:ring-indigo-500/10"
              />
              {searchTerm && (
                <button
                  onClick={() => setSearchTerm("")}
                  aria-label="Hapus pencarian"
                  className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400 transition hover:text-slate-600"
                >
                  <X className="h-4 w-4" />
                </button>
              )}
            </div>

            {/* <button
              onClick={exportToExcel}
              disabled={filteredData.length === 0}
              className="flex shrink-0 items-center justify-center gap-2 rounded-sm bg-gradient-to-r from-indigo-600 to-blue-600 px-5 py-3 text-sm font-semibold text-white shadow-md shadow-indigo-500/25 transition hover:from-indigo-700 hover:to-blue-700 disabled:cursor-not-allowed disabled:from-slate-300 disabled:to-slate-300 disabled:shadow-none"
            >
              <Download className="h-4 w-4" />
              Unduh Excel
            </button> */}
          </div>

          {/* Filter Status Progress — pill berwarna, berdasarkan kolom PROGRESS PER 24 AGUSTUS 2026 */}
          <div className="flex flex-wrap items-center justify-between gap-3 border-t border-slate-100 pt-4">
            <div className="flex flex-wrap gap-2">
              {FILTER_TABS.map((tab) => (
                <FilterPill
                  key={tab.key}
                  active={statusFilter === tab.key}
                  onClick={() => setStatusFilter(tab.key)}
                  colorClass={PILL_GRADIENT[tab.key]}
                >
                  {tab.label}
                  {tab.key !== "ALL" && (
                    <span className="ml-1.5 opacity-80">
                      ({statusCounts[tab.key as StatusKey]})
                    </span>
                  )}
                </FilterPill>
              ))}
            </div>

            {isFiltering && (
              <button
                onClick={resetFilter}
                className="text-xs font-semibold text-rose-500 transition hover:text-rose-600 hover:underline"
              >
                Reset Filter
              </button>
            )}
          </div>
        </div>

        {/* Tabel Data */}
        <div className="overflow-hidden rounded-xl border border-slate-200/70 bg-white shadow-[0_1px_2px_rgba(15,23,42,0.04),0_18px_40px_-20px_rgba(15,23,42,0.16)]">
          <div
            ref={scrollRef}
            onScroll={handleTableScroll}
            className="overflow-x-auto overscroll-x-contain [-webkit-overflow-scrolling:touch] [&::-webkit-scrollbar]:hidden [scrollbar-width:none]"
          >
            <table className="w-full min-w-[1150px] border-collapse text-left text-xs">
              <thead>
                <tr className="bg-gradient-to-r from-indigo-50 via-blue-50 to-sky-50 text-[11px] font-extrabold uppercase tracking-wide text-indigo-800">
                  <th className="whitespace-nowrap border-b-2 border-indigo-100 px-4 py-3.5 text-left">
                    No
                  </th>
                  <th className="whitespace-nowrap border-b-2 border-indigo-100 px-4 py-3.5 text-left">
                    Area Intervensi
                  </th>
                  <th className="whitespace-nowrap border-b-2 border-indigo-100 px-4 py-3.5 text-left">
                    OPD
                  </th>
                  <th className="border-b-2 border-indigo-100 px-4 py-3.5 text-left">
                    Data / Dokumen yang Diperlukan
                  </th>
                  <th className="border-b-2 border-indigo-100 px-4 py-3.5 text-left">
                    Komitmen Berita Acara
                    <br />
                    13 Agustus 2026
                  </th>
                  <th className="border-b-2 border-indigo-100 px-4 py-3.5 text-left">
                    Progress per 24 Agustus 2026
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredData.length === 0 ? (
                  <tr>
                    <td colSpan={6} className="p-12 text-center">
                      <div className="mx-auto mb-2 flex h-12 w-12 items-center justify-center rounded-xl bg-amber-50 text-amber-500">
                        <Info className="h-6 w-6" />
                      </div>
                      <p className="text-sm font-semibold text-slate-700">
                        Data Tidak Ditemukan
                      </p>
                      <p className="mt-0.5 text-xs text-slate-400">
                        Tidak ada dokumen yang cocok dengan pencarian / filter saat ini.
                      </p>
                    </td>
                  </tr>
                ) : (
                  filteredData.map((row, idx) => {
                    const kategori = classifyStatus(row.progress_per_24_agustus_2026);
                    const cfg = STATUS_CONFIG[kategori];
                    return (
                      <tr
                        key={`${row.no}-${idx}`}
                        className="transition-colors odd:bg-white even:bg-slate-50/60 hover:bg-indigo-50/50"
                      >
                        <td className="whitespace-nowrap px-4 py-3 align-top font-semibold text-slate-500">
                          {row.no}
                        </td>
                        <td className="max-w-[200px] px-4 py-3 align-top font-medium text-indigo-700">
                          {row.area_intervensi}
                        </td>
                        <td className="max-w-[180px] px-4 py-3 align-top text-slate-600">
                          {row.opd}
                        </td>
                        <td className="max-w-[320px] px-4 py-3 align-top whitespace-pre-line text-slate-700">
                          {row.data_dokumen_yg_diperlukan}
                        </td>
                        <td className="max-w-[220px] px-4 py-3 align-top whitespace-pre-line text-slate-600">
                          {row.komitmen_berita_acara_13_agustus_2026}
                        </td>
                        <td className="max-w-[260px] px-4 py-3 align-top">
                          <span
                            className={`inline-flex items-center gap-1.5 whitespace-pre-line rounded-lg border px-2.5 py-1 text-[11px] font-semibold ${cfg.badge}`}
                          >
                            <span className={`h-1.5 w-1.5 shrink-0 rounded-full ${cfg.dot}`} />
                            {row.progress_per_24_agustus_2026}
                          </span>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* Indikator scroll horizontal buatan */}
          <div className="relative h-1.5 w-full overflow-hidden bg-slate-100">
            <div
              className="absolute inset-y-0 rounded-xl bg-gradient-to-r from-indigo-500 to-sky-400 transition-[left,width] duration-100"
              style={{
                width: `${scrollBar.widthPct}%`,
                left: `${scrollBar.leftPct}%`,
              }}
            />
          </div>

          <div className="flex items-center justify-between border-t border-slate-100 bg-slate-50/70 px-6 py-3 text-xs text-slate-500">
            <span>
              Menampilkan{" "}
              <strong className="font-semibold text-indigo-700">{filteredData.length}</strong>{" "}
              dari <strong className="font-semibold text-slate-700">{data.length}</strong> total
              item dokumen
            </span>
          </div>
        </div>
      </div>
    </main>
  );
}

function StatCard({
  icon,
  gradient,
  label,
  value,
  active,
  onClick,
}: {
  icon: React.ReactNode;
  gradient: string;
  label: string;
  value: React.ReactNode;
  active?: boolean;
  onClick?: () => void;
}) {
  return (
    <button
      onClick={onClick}
      className={`flex items-center gap-3 rounded-xl border bg-white p-4 text-left shadow-[0_1px_2px_rgba(15,23,42,0.04),0_10px_26px_-16px_rgba(15,23,42,0.14)] transition hover:-translate-y-0.5 hover:shadow-[0_1px_2px_rgba(15,23,42,0.04),0_16px_32px_-14px_rgba(79,70,229,0.20)] ${
        active ? "border-indigo-300 ring-2 ring-indigo-500/15" : "border-slate-200/70"
      }`}
    >
      <div
        className={`flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br text-white shadow-md ${gradient}`}
      >
        {icon}
      </div>
      <div className="min-w-0">
        <p className="truncate text-[11px] font-medium text-slate-500">{label}</p>
        <p className="text-xl font-extrabold text-slate-900">{value}</p>
      </div>
    </button>
  );
}

function FilterPill({
  active,
  onClick,
  colorClass,
  children,
}: {
  active: boolean;
  onClick: () => void;
  colorClass: string;
  children: React.ReactNode;
}) {
  return (
    <button
      onClick={onClick}
      className={`cursor-pointer rounded-lg px-4 py-2 text-xs font-semibold transition ${
        active
          ? `bg-gradient-to-r text-white shadow-md ${colorClass}`
          : "bg-slate-100 text-slate-600 hover:bg-slate-200"
      }`}
    >
      {children}
    </button>
  );
}
