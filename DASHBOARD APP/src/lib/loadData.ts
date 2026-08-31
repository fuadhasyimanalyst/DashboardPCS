import * as XLSX from 'xlsx';
import type { SalesRow, TargetRow, UangMasukRow } from './types';
import { deriveDateParts, normalizeBulanToMonthNum } from './types';

// ---------------------------------------------------------------------------
// MODE LOKAL: dashboard membaca 3 file Excel langsung di browser (SheetJS),
// dari public/data/. Ini mode yang didokumentasikan di README ("Status saat
// ini: mode LOKAL"). Tidak butuh Supabase/Vercel/.env sama sekali.
//
// Kode migrasi ke Supabase (endpoint /api/meta & /api/data) sudah disiapkan
// di api/meta.js dan api/data.js untuk dipakai NANTI kalau mau pindah mode
// (lihat bagian "Migrasi ke Supabase" di README) -- tapi endpoint itu HANYA
// jalan di atas Vercel Serverless Functions, bukan di `npm run dev` biasa,
// jadi tidak dipakai di sini.

const SALES_URL = '/data/DATA.xlsx';
const TARGET_URL = '/data/DATA_TARGET_FUAD.xlsx';
const UANG_MASUK_URL = '/data/REALISASI%20UANG%20MASUK.xlsx';

interface WorkbookCacheEntry {
  sales: Record<string, unknown>[];
  targets: Record<string, unknown>[];
  uangMasuk: Record<string, unknown>[];
}

// Cache di memori: ketiga file cuma diambil & di-parse SEKALI per sesi
// (sampai tombol Refresh memanggil resetDataCache()), lalu dipakai bareng
// oleh loadSalesData/loadTargetData/loadUangMasukData.
let payloadPromise: Promise<WorkbookCacheEntry> | null = null;

function normalizeKey(k: string): string {
  return k.replace(/^["'\s]+|["'\s]+$/g, '').trim().toUpperCase();
}

function normalizeRow(row: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(row)) out[normalizeKey(k)] = v;
  return out;
}

async function readSheet(url: string): Promise<Record<string, unknown>[]> {
  // cache: 'no-store' + cache-busting query -> tombol Refresh (lihat
  // resetDataCache) selalu membaca file terbaru dari disk, bukan versi
  // basi yang sempat disimpan browser.
  const res = await fetch(`${url}?t=${Date.now()}`, { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`Gagal memuat ${url} (status ${res.status})`);
  }
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: '' });
}

function toNumber(v: unknown): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    const cleaned = v.replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.');
    const n = parseFloat(cleaned);
    return isNaN(n) ? 0 : n;
  }
  return 0;
}

// TGL FAKTUR bisa berupa objek Date (kalau cellDates:true berhasil mendeteksi
// format tanggal sel), serial Excel, atau string "dd/mm/yyyy" -- sama seperti
// scripts/sync-data.mjs, tapi versi browser (tanpa Node.js) dan mengembalikan
// Date lokal langsung (bukan string ISO), karena dashboard butuh Date lokal.
function parseExcelDate(v: unknown): Date | null {
  if (v instanceof Date && !isNaN(v.getTime())) {
    // Bulatkan ke menit terdekat untuk hilangkan noise floating-point serial
    // Excel (mis. 23:59:56 -> seharusnya tengah malam hari berikutnya).
    const rounded = new Date(Math.round(v.getTime() / 60000) * 60000);
    return new Date(rounded.getFullYear(), rounded.getMonth(), rounded.getDate());
  }
  if (typeof v === 'number') {
    // Serial tanggal Excel (hari sejak 1899-12-30 versi Windows).
    const utcMs = Math.round((v - 25569) * 86400 * 1000);
    const d = new Date(utcMs);
    if (isNaN(d.getTime())) return null;
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }
  if (typeof v === 'string' && v.trim()) {
    const m = v.trim().match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
    if (m) {
      const [, dd, mm, yyRaw] = m;
      const yyyy = yyRaw.length === 2 ? 2000 + Number(yyRaw) : Number(yyRaw);
      const d = new Date(yyyy, Number(mm) - 1, Number(dd));
      return isNaN(d.getTime()) ? null : d;
    }
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

async function fetchPayload(): Promise<WorkbookCacheEntry> {
  const [sales, targets, uangMasuk] = await Promise.all([
    readSheet(SALES_URL),
    readSheet(TARGET_URL),
    // Uang masuk bersifat opsional -- kalau file belum ada/gagal dibaca,
    // jangan gagalkan seluruh dashboard.
    readSheet(UANG_MASUK_URL).catch(() => []),
  ]);
  return { sales, targets, uangMasuk };
}

function getPayload(): Promise<WorkbookCacheEntry> {
  if (!payloadPromise) {
    payloadPromise = fetchPayload().catch((err) => {
      payloadPromise = null;
      throw err;
    });
  }
  return payloadPromise;
}

/** Dipanggil oleh tombol Refresh supaya percobaan berikutnya membaca ulang file Excel dari disk. */
export function resetDataCache() {
  payloadPromise = null;
}

export async function loadSalesData(): Promise<SalesRow[]> {
  const { sales: raw } = await getPayload();
  return raw
    .map((r): SalesRow => {
      const row = normalizeRow(r);
      const tglFaktur = parseExcelDate(row['TGL FAKTUR']);
      // Fallback untuk file DATA.xlsx format LAMA (belum punya kolom TGL
      // FAKTUR, hanya BULAN + TAHUN terpisah): kalau TGL FAKTUR kosong/tidak
      // valid, pakai BULAN/TAHUN mentah dari file langsung. Tanggal harian
      // (tanggal/tanggalStr) tidak bisa diketahui dari format lama ini, jadi
      // dibiarkan kosong -- hanya memengaruhi halaman "Omset Harian" yang
      // memang butuh data per-tanggal.
      const parts = tglFaktur
        ? deriveDateParts(tglFaktur)
        : {
            bulan: String(row['BULAN'] ?? '').trim(),
            monthNum: normalizeBulanToMonthNum(row['BULAN'] as string | number | undefined | null),
            tahun: row['TAHUN'] ? Number(row['TAHUN']) : new Date().getFullYear(),
            tanggal: 0,
            tanggalStr: '',
          };
      const { bulan, monthNum, tahun, tanggal, tanggalStr } = parts;
      return {
        noFaktur: String(row['NO FAKTUR'] ?? '').trim(),
        nominal: toNumber(row['NOMINAL']),
        supp: String(row['SUPP'] ?? '').trim(),
        depo: String(row['DEPO'] ?? '').trim().toUpperCase(),
        tglFaktur,
        tanggalStr,
        tanggal,
        bulan,
        monthNum,
        tahun,
        kdGrup: String(row['KD GRUP'] ?? '').trim(),
        sales: String(row['SALES'] ?? '').trim(),
        kota: String(row['KOTA'] ?? '').trim().toUpperCase(),
        kecamatan: String(row['KECAMATAN'] ?? '').trim(),
        tele: String(row['TELE'] ?? '').trim(),
        kodePelanggan: String(row['KODE PELANGGAN'] ?? '').trim(),
        namaPelanggan: String(row['NAMA PELANGGAN'] ?? '').trim(),
        alamatPelanggan: String(row['ALAMAT PELANGGAN'] ?? '').trim(),
        namaBarang: String(row['NAMA BARANG'] ?? '').trim(),
        qty: toNumber(row['QTY']),
        rankBayar: String(row['RANK BAYAR'] ?? '').trim(),
        rankOmset: String(row['RANK OMSET'] ?? '').trim(),
      };
    })
    .filter((r) => r.depo && r.sales);
}

// Mode lokal tidak punya konsep "waktu sinkronisasi server" (itu cuma ada di
// mode Supabase) -- TopBar cukup menyembunyikan baris ini kalau null.
export async function loadDataSyncedAt(): Promise<Date | null> {
  return null;
}

export async function loadUangMasukData(): Promise<UangMasukRow[]> {
  const { uangMasuk: raw } = await getPayload();
  return raw
    .map((r): UangMasukRow => {
      const row = normalizeRow(r);
      const bulanRaw = String(row['BULAN'] ?? '').trim();
      return {
        tahun: row['TAHUN'] ? Number(row['TAHUN']) : new Date().getFullYear(),
        bulan: bulanRaw,
        monthNum: normalizeBulanToMonthNum(bulanRaw),
        depo: String(row['DEPO'] ?? '').trim().toUpperCase(),
        targetPiutang: toNumber(row['TARGET PIUTANG']),
        realisasiPiutang: toNumber(row['REALISASI PIUTANG']),
      };
    })
    .filter((r) => r.depo);
}

export async function loadTargetData(): Promise<TargetRow[]> {
  const { targets: raw } = await getPayload();
  const monthCols = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  return raw
    .map((r): TargetRow => {
      const row = normalizeRow(r);
      return {
        namaSalesman: String(row['NAMA SALESMAN'] ?? '').trim(),
        depo: String(row['DEPO'] ?? '').trim().toUpperCase(),
        supplier: String(row['SUPPLIER'] ?? '').trim().toUpperCase(),
        tahun: row['TAHUN'] ? Number(row['TAHUN']) : new Date().getFullYear(),
        monthly: monthCols.map((c) => toNumber(row[c])),
      };
    })
    .filter((r) => r.namaSalesman);
}
