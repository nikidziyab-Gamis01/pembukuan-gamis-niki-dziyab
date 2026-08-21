import { db, auth, googleProvider } from "./firebase";
import { signInWithPopup, signOut, onAuthStateChanged } from "firebase/auth";
import { doc, getDoc, onSnapshot, setDoc } from "firebase/firestore";
import { useState, useEffect, useMemo, useRef } from "react";
import {
  LayoutDashboard,
  Shirt,
  ArrowDownToLine,
  ArrowUpFromLine,
  RotateCcw,
  XCircle,
  Wallet,
  Settings2,
  Plus,
  Trash2,
  Search,
  AlertTriangle,
  ChevronDown,
  Save,
  Download,
  Upload,
  FileSpreadsheet,
  FileText,
  Printer,
  ChevronUp,
  PieChart,
} from "lucide-react";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Cell,
} from "recharts";
import * as XLSX from "xlsx";

// ---------- PREMIUM COLOR SCHEME (Niki Dziyab Theme) ----------
const C = {
  primary: "#7c2d12",
  secondary: "#9a3412",
  accent: "#10b981",
  bg: "#fcfaf8",
  text: "#1f2937",
  surface: "#FFFFFF",
  surfaceAlt: "#F3ECE1",
  primaryDark: "#54202A",
  primarySoft: "#F1E1E1",
  accentSoft: "#F3E4C2",
  muted: "#8C7B6B",
  border: "#E7DBCA",
  danger: "#B3432B",
  success: "#4B7A51",
};

const WARNA_SWATCH = {
  HITAM: "#1B1B1B",
  MAHOGANI: "#5A3324",
  EMRALD: "#0F6E4F",
  MATCHIATO: "#6F4E37",
  BURGUNDY: "#7A2E3B",
};

const BULAN_ID = [
  "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
  "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER",
];

const TOKO_OPTIONS = ["Shopee", "TikTok", "Tokopedia", "Lazada", "Website", "Offline"];
const PENGIRIMAN_OPTIONS = ["Shopee", "TikTokShop", "Tokopedia", "Lazada", "Manual"];
const ONGKIR_OPTIONS = ["Potong di Aplikasi", "Manual"];

const EMPTY_MASTER = { vendor: [], jenis: [], brand: [], bahan: [], warna: [], ukuran: [] };

const normalizeData = (value) => {
  const source = value && typeof value === "object" ? value : {};
  const master = source.masterCode && typeof source.masterCode === "object" ? source.masterCode : {};

  return {
    products: Array.isArray(source.products) ? source.products : [],
    masuk: Array.isArray(source.masuk) ? source.masuk : [],
    keluar: Array.isArray(source.keluar) ? source.keluar : [],
    retur: Array.isArray(source.retur) ? source.retur : [],
    reject: Array.isArray(source.reject) ? source.reject : [],
    masterCode: {
      vendor: Array.isArray(master.vendor) ? master.vendor : EMPTY_MASTER.vendor,
      jenis: Array.isArray(master.jenis) ? master.jenis : EMPTY_MASTER.jenis,
      brand: Array.isArray(master.brand) ? master.brand : EMPTY_MASTER.brand,
      bahan: Array.isArray(master.bahan) ? master.bahan : EMPTY_MASTER.bahan,
      warna: Array.isArray(master.warna) ? master.warna : EMPTY_MASTER.warna,
      ukuran: Array.isArray(master.ukuran) ? master.ukuran : EMPTY_MASTER.ukuran,
    },
    kasbon: Array.isArray(source.kasbon) ? source.kasbon : [],
    karyawanMaster: Array.isArray(source.karyawanMaster) ? source.karyawanMaster : [],
    bukuKas: Array.isArray(source.bukuKas) ? source.bukuKas : [],
    gajiSetelan: {
      gajiPokokOwner: Number(source.gajiSetelan?.gajiPokokOwner) || 0,
      komisiOwnerPerPcs: Number(source.gajiSetelan?.komisiOwnerPerPcs) || 0,
      komisiHostPerPcs: Number(source.gajiSetelan?.komisiHostPerPcs) || 0,
      komisiAdminPerPcs: Number(source.gajiSetelan?.komisiAdminPerPcs) || 0,
    },
    gajiDinamis: source.gajiDinamis && typeof source.gajiDinamis === "object" ? source.gajiDinamis : {},
  };
};

const EMPTY_DATA = normalizeData({
  products: [], masuk: [], keluar: [], retur: [], reject: [], masterCode: EMPTY_MASTER, kasbon: [], karyawanMaster: [], bukuKas: []
});

const todayStr = () => new Date().toISOString().slice(0, 10);
const monthOf = (dateStr) => {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (isNaN(d)) return "";
  return BULAN_ID[d.getMonth()];
};
const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
const fmt = (n) => (Number(n) || 0).toLocaleString("id-ID");

function downloadBlob(content, filename, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function buildRows(data, stockMap, stokAkhir) {
  if (!data) return { produk: [], masuk: [], keluar: [], retur: [], reject: [], masterCode: [], pencairan: [] };

  const produk = (data.products || []).map((p) => ({
    Kode: p.kode, "Nama Barang": p.nama, Warna: p.warna, Ukuran: p.ukuran, Vendor: p.vendor, "HPP Dasar": p.hpp || 0, "Stok Akhir": stokAkhir(p.kode)
  }));

  const masuk = (data.masuk || []).map((m) => {
    const matchedProd = (data.products || []).find((p) => p.kode === m.kode) || {};
    return {
      Tanggal: m.tanggal, "Kode Barang": m.kode, "Nama Barang": matchedProd.nama || "—", Qty: m.qty, Vendor: m.vendor, "HPP Satuan": m.hpp || matchedProd.hpp || 0, "Total Nilai HPP": (m.qty * (m.hpp || matchedProd.hpp || 0)), Catatan: m.catatan
    };
  });

  const keluar = (data.keluar || []).map((k) => {
    const matchedProd = (data.products || []).find((p) => p.kode === k.kode) || {};
    return {
      Tanggal: k.tanggal, "Kode Pesanan": k.kodePesanan, Shift: k.shift, Host: k.host,
      "Kode Barang": k.kode, "Nama Barang": matchedProd.nama || "—", Qty: k.qty, Toko: k.toko, "Harga Satuan Aplikasi": k.hargaJual || 0, "Total Omset Bruto": (k.qty * (k.hargaJual || 0)), "Nominal Cair Riil": k.nominalCairRiil || 0, "COD/NON COD": k.cod,
      Affiliate: k.affiliate ? "Ya" : "Tidak", "Nama Affiliate": k.namaAffiliate || "",
      Pengiriman: k.pengiriman || "", Ongkir: k.ongkir || "", Catatan: k.catatan || "", "Sudah Cair": k.cair ? "Ya" : "Belum",
    };
  });

  const retur = (data.retur || []).map((r) => ({
    Tanggal: r.tanggal, "Kode Pesanan": r.kodePesanan, "Kode Barang": r.kode, "Kode Retur": r.kodeRetur, Qty: r.qty, "Sesuai/Tidak": r.sesuai, Ekspedisi: r.ekspedisi, Catatan: r.catatan,
  }));

  const reject = (data.reject || []).map((r) => ({
    Tanggal: r.tanggal, "Kode Barang": r.kode, "Nama Barang": ((data.products || []).find((p) => p.kode === r.kode) || {}).nama, Qty: r.qty, Toko: r.toko, Catatan: r.catatan,
  }));

  const masterCode = [];
  const targetMaster = data.masterCode || {};
  Object.entries(targetMaster).forEach(([kategori, list]) => {
    if (Array.isArray(list)) {
      list.forEach((item) => {
        masterCode.push({ Kategori: kategori.toUpperCase(), Kode: item.code, Nama: item.name });
      });
    }
  });

  const totalKeluar = (data.keluar || []).reduce((s, k) => s + (Number(k.qty) || 0), 0);
  const totalCair = (data.keluar || []).filter((k) => k.cair).reduce((s, k) => s + (Number(k.qty) || 0), 0);
  const pencairan = [{ "Barang Keluar": totalKeluar, "Sudah Cair": totalCair, Selisih: totalKeluar - totalCair }];
  
  return { produk, masuk, keluar, retur, reject, masterCode, pencairan };
}

function exportExcel(data, stockMap, stokAkhir) {
  const { produk, masuk, keluar, retur, reject, masterCode, pencairan } = buildRows(data, stockMap, stokAkhir);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(produk), "Produk & Stok");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(masuk), "Barang Masuk");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(keluar), "Barang Keluar");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(retur), "Retur");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(reject), "Reject");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pencairan), "Pencairan");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(masterCode), "Kode Master");
  XLSX.writeFile(wb, `Pembukuan-Gamis-Niki-Dziyab-${todayStr()}.xlsx`);
}

function tableHtml(title, rows) {
  if (!rows.length) return `<h3>${title}</h3><p style="color:#8C7B6B;font-size:12px;">Belum ada data.</p>`;
  const headers = Object.keys(rows[0]);
  const thead = headers.map((h) => `<th style="border:1px solid #ccc;padding:5px 8px;background:#F1E1E1;text-align:left;">${h}</th>`).join("");
  const tbody = rows.map((r) => `<tr>${headers.map((h) => `<td style="border:1px solid #ccc;padding:5px 8px;">${r[h] ?? ""}</td>`).join("")}</tr>`).join("");
  return `<h3 style="font-family:Georgia,serif;color:#54202A;margin:22px 0 8px;">${title}</h3>
    <table style="border-collapse:collapse;width:100%;font-size:11px;font-family:Arial,sans-serif;"><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table>`;
}

function exportWord(data, stockMap, stokAkhir) {
  const { produk, masuk, keluar, retur, reject, masterCode, pencairan } = buildRows(data, stockMap, stokAkhir);
  const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://w3.org">
  <head><meta charset="utf-8"><title>Pembukuan Gamis Niki Dziyab</title></head>
  <body style="font-family:Arial,sans-serif;">
    <h1 style="font-family:Georgia,serif;color:#54202A;margin-bottom:0;">Pembukuan Gamis Niki Dziyab</h1>
    <p style="color:#8C7B6B;font-size:12px;margin-top:2px;">Dicetak: ${todayStr()}</p>
    ${tableHtml("Produk & Stok", produk)}
    ${tableHtml("Barang Masuk", masuk)}
    ${tableHtml("Barang Keluar", keluar)}
    ${tableHtml("Retur", retur)}
    ${tableHtml("Reject", reject)}
    ${tableHtml("Pencairan", pencairan)}
    ${tableHtml("Kode Master", masterCode)}
  </body></html>`;
  downloadBlob(html, `Pembukuan-Gamis-Niki-Dziyab-${todayStr()}.doc`, "application/msword");
}

const TABS = [
  { id: "dashboard", label: "Dasbor", icon: LayoutDashboard },
  { id: "produk", label: "Produk & Stok", icon: Shirt },
  { id: "masuk", label: "Barang Masuk", icon: ArrowDownToLine },
  { id: "keluar", label: "Barang Keluar", icon: ArrowUpFromLine },
  { id: "retur", label: "Retur", icon: RotateCcw },
  { id: "reject", label: "Reject", icon: XCircle },
  { id: "pencairan", label: "Pencairan", icon: Wallet },
  { id: "bukuKas", label: "📊 Master Buku Kas", icon: PieChart },
  { id: "keuangan", label: "💰 Keuangan & Gaji Tim", icon: Settings2 },
  { id: "master", label: "⚙️ Kode Master", icon: Settings2 },
];

// ---------- HELPER TRANSAKSI & SLIP GAJI TIM ----------
function calculateSalarySlips(data) {
  const safeData = normalizeData(data || {});
  const slips = [];
  const employeeMap = {};

  // Kumpulkan semua karyawan dari karyawanMaster resmi
  (safeData.karyawanMaster || []).forEach((emp) => {
    const namaEmp = String(emp?.nama || "").trim();
    if (!namaEmp) return;
    
    employeeMap[namaEmp] = {
      nama: namaEmp,
      role: emp?.peran || "Tim Live",
      qtyKeluar: 0,
      gajiPokok: 0,
      komisiPerPcs: 0,
    };
  });

  // Hitung performa qty kuantitas jualan live/shift per karyawan
  (safeData.keluar || []).forEach((k) => {
    const hostName = String(k?.host || "").trim();
    const adminName = String(k?.shift || "").trim();
    const qty = Number(k?.qty) || 0;

    if (hostName && employeeMap[hostName]) {
      employeeMap[hostName].qtyKeluar += qty;
    }
    if (adminName && adminName !== hostName && employeeMap[adminName]) {
      employeeMap[adminName].qtyKeluar += qty;
    }
  });

  // Pemetaan total kasbon berjalan per karyawan
  const kasbonMap = {};
  (safeData.kasbon || []).forEach((k) => {
    const nama = String(k?.nama || "").trim();
    if (!kasbonMap[nama]) kasbonMap[nama] = 0;
    kasbonMap[nama] += Number(k?.nominal) || 0;
  });

  // Ambil data konfigurasi gaji dinamis manual
  const gajiDinamis = data?.gajiDinamis || {};

  // Kompilasi output slip tim operasional
  Object.values(employeeMap).forEach((emp) => {
    const gajiPokokEmp = (gajiDinamis?.[emp.nama]?.gajiPokok) || 0;
    const komisiPerPcsEmp = (gajiDinamis?.[emp.nama]?.komisiPerPcs) || 0;
    const bruto = gajiPokokEmp + emp.qtyKeluar * komisiPerPcsEmp;
    const kasbonTotal = kasbonMap[emp.nama] || 0;
    const bersih = bruto - kasbonTotal;

    slips.push({
      nama: emp.nama,
      role: emp.role,
      qtyKeluar: emp.qtyKeluar,
      bruto,
      kasbon: kasbonTotal,
      bersih: Math.max(0, bersih),
    });
  });

  return slips.sort((a, b) => (a?.nama || "").localeCompare(b?.nama || ""));
}

function calculateTotalSalaryLoad(data) {
  const slips = calculateSalarySlips(data);
  return (slips || []).reduce((total, slip) => total + (slip?.bruto || 0), 0);
}
// ---------- FUNGSI PRINTER BLUETOOTH THERMAL ----------
async function printViaBluetooth(item, jenis) {
  try {
    const device = await navigator.bluetooth.requestDevice({
      filters: [{ services: ["00001101-0000-1000-8000-00805f9b34fb"] }],
    });

    const server = await device.gatt.connect();
    const service = await server.getPrimaryService("00001101-0000-1000-8000-00805f9b34fb");
    const characteristic = await service.getCharacteristic("00002a19-0000-1000-8000-00805f9b34fb");

    const receiptText = `
      ========= NIKI DZIYAB ==========
      Tipe: ${jenis === "keluar" ? "BARANG KELUAR" : "TRANSAKSI"}
      Tanggal: ${todayStr()}
      Jam: ${new Date().toLocaleTimeString("id-ID")}
      ===================================
      Kode SKU: ${item?.kode || "—"}
      Qty: ${item?.qty || 0} Pcs
      Pesanan: ${item?.kodePesanan || "—"}
      Platform: ${item?.toko || "—"}
      Nominal Jual: Rp ${(item?.hargaJual || 0).toLocaleString("id-ID")}
      ===================================
      Cetak Otomatis Sistem Niki Dziyab
      ===================================
    `;

    const encoder = new TextEncoder();
    const data = encoder.encode(receiptText);
    await characteristic.writeValue(data);

    alert("Struk berhasil dicetak ke printer Bluetooth!");
  } catch (error) {
    console.error("Error Bluetooth Printer:", error);
    alert("Gagal terhubung ke printer Bluetooth. Pastikan bluetooth aktif dan printer dalam jangkauan pairing.");
  }
}

// ---------- MASTER BUKU KAS (DENGAN SALDO BERJALAN KRONOLOGIS) ----------
function BukuKasTab({ data, save, flash }) {
  const safeData = normalizeData(data || {});
  const [formJurnal, setFormJurnal] = useState({
    tanggal: todayStr(),
    jenis: "Debit",
    nominal: "",
    keterangan: "",
  });

  const submitJurnal = () => {
    if (!Number(formJurnal.nominal)) {
      flash("Masukkan nominal uang!");
      return;
    }
    const entry = {
      id: uid(),
      tanggal: formJurnal.tanggal,
      jenis: formJurnal.jenis,
      nominal: Number(formJurnal.nominal),
      keterangan: formJurnal.keterangan,
    };
    save({
      ...safeData,
      bukuKas: [entry, ...(safeData.bukuKas || [])],
    });
    setFormJurnal({ tanggal: todayStr(), jenis: "Debit", nominal: "", keterangan: "" });
    flash("Jurnal manual kas toko berhasil dicatat!");
  };

  const removeJurnal = (id) => {
    if (confirm("Apakah Anda yakin ingin menghapus data jurnal kas ini?")) {
      save({ ...safeData, bukuKas: (safeData.bukuKas || []).filter((b) => b.id !== id) });
      flash("Data jurnal kas dihapus.");
    }
  };

  // 1. Hitung total nominal Debit (Masuk) dan Kredit (Keluar)
  const totalDebit = (safeData.bukuKas || [])
    .filter((b) => b.jenis === "Debit")
    .reduce((s, b) => s + (Number(b.nominal) || 0), 0);

  const totalKredit = (safeData.bukuKas || [])
    .filter((b) => b.jenis === "Kredit")
    .reduce((s, b) => s + (Number(b.nominal) || 0), 0);

  const saldoAkhir = totalDebit - totalKredit;

  // 2. Sorting data berdasarkan tanggal dari TERLAMA ke TERBARU untuk hitung saldo akumulatif kumulatif
  const sortedKronologis = [...(safeData.bukuKas || [])].sort(
    (a, b) => new Date(a.tanggal) - new Date(b.tanggal)
  );

  let akumulasiSaldo = 0;
  const kronologisWithSaldo = sortedKronologis.map((b) => {
    const nominalItem = Number(b.nominal) || 0;
    if (b.jenis === "Debit") {
      akumulasiSaldo += nominalItem;
    } else {
      akumulasiSaldo -= nominalItem;
    }
    return {
      ...b,
      saldoKumulatif: akumulasiSaldo,
    };
  });

  // 3. Membalikkan kembali urutan agar data transaksi TERBARU berada di baris paling atas tabel
  const bukuKasWithSaldo = [...kronologisWithSaldo].reverse();

  return (
    <div>
      {/* KOTAK REKAP MUTASI AKUNTANSI */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 12, marginBottom: 20 }}>
        <div className="gm-card" style={{ padding: 18 }}>
          <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600, textTransform: "uppercase", letterSpacing: ".03em" }}>Total Uang Cair Masuk (Debit)</div>
          <div className="gm-h1" style={{ fontSize: 24, fontWeight: 700, color: C.success, marginTop: 8 }}>Rp {fmt(totalDebit)}</div>
        </div>
        <div className="gm-card" style={{ padding: 18 }}>
          <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600, textTransform: "uppercase", letterSpacing: ".03em" }}>Total Uang Keluar (Kredit)</div>
          <div className="gm-h1" style={{ fontSize: 24, fontWeight: 700, color: C.danger, marginTop: 8 }}>Rp {fmt(totalKredit)}</div>
        </div>
        <div className="gm-card" style={{ padding: 18 }}>
          <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600, textTransform: "uppercase", letterSpacing: ".03em" }}>Saldo Netto Buku Kas Aktif</div>
          <div className="gm-h1" style={{ fontSize: 24, fontWeight: 700, color: saldoAkhir >= 0 ? C.accent : C.danger, marginTop: 8 }}>Rp {fmt(saldoAkhir)}</div>
        </div>
      </div>

      {/* FORM INPUT MANUAL ARUS JURNAL NON-PLATFORM */}
      <SectionCard title="Input Manual Jurnal Buku Kas" subtitle="Gunakan modul ini untuk mencatatkan pengeluaran operasional di luar transaksi market seperti token listrik, air, atau sewa toko">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: 12, alignItems: "end" }}>
          <Field label="Tanggal"><input className="gm-input" type="date" value={formJurnal.tanggal} onChange={(e) => setFormJurnal({ ...formJurnal, tanggal: e.target.value })} /></Field>
          <Field label="Jenis Arus Kas">
            <select className="gm-select" value={formJurnal.jenis} onChange={(e) => setFormJurnal({ ...formJurnal, jenis: e.target.value })}>
              <option value="Debit">Debit (Uang Masuk / Modal)</option>
              <option value="Kredit">Kredit (Uang Keluar / Biaya)</option>
            </select>
          </Field>
          <Field label="Nominal Uang (Rp)"><input className="gm-input" type="number" min="0" value={formJurnal.nominal} onChange={(e) => setFormJurnal({ ...formJurnal, nominal: e.target.value })} placeholder="0" /></Field>
          <Field label="Keterangan Nota"><input className="gm-input" value={formJurnal.keterangan} onChange={(e) => setFormJurnal({ ...formJurnal, keterangan: e.target.value })} placeholder="Biaya lakban / bensin packing" /></Field>
          <button className="gm-btn" onClick={submitJurnal}><Plus size={15} /> Catat Jurnal</button>
        </div>
      </SectionCard>

      {/* REKENING KORAN MUTASI KAS UTUH */}
      <SectionCard title="Rekening Koran & Saldo Berjalan" subtitle="Daftar rekam jejak mutasi aliran dana masuk dan keluar toko secara kronologis">
        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th>No</th><th>Tanggal</th><th>Deskripsi Alur Transaksi</th><th>Debit (+)</th><th>Kredit (-)</th><th>Saldo Berjalan Kumulatif</th><th>Aksi</th>
              </tr>
            </thead>
            <tbody>
              {bukuKasWithSaldo.length === 0 && <EmptyRow colSpan={7} text="Belum ada histori transaksi kas." />}
              {bukuKasWithSaldo.map((b, idx) => (
                <tr key={b.id}>
                  <td style={{ textAlign: "center", color: C.muted, fontSize: 12 }}>{idx + 1}</td>
                  <td>{b.tanggal}</td>
                  <td style={{ whiteSpace: "normal", minWidth: 220, fontSize: 12.5 }}>{b.keterangan}</td>
                  <td style={{ fontWeight: 700, color: b.jenis === "Debit" ? C.success : "inherit" }}>
                    {b.jenis === "Debit" ? `Rp ${fmt(b.nominal || 0)}` : "—"}
                  </td>
                  <td style={{ fontWeight: 700, color: b.jenis === "Kredit" ? C.danger : "inherit" }}>
                    {b.jenis === "Kredit" ? `Rp ${fmt(b.nominal || 0)}` : "—"}
                  </td>
                  <td style={{ fontWeight: 700, color: C.primary, fontSize: 12.5 }}>Rp {fmt(b.saldoKumulatif || 0)}</td>
                  <td><DeleteBtn onClick={() => removeJurnal(b.id)} /></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </SectionCard>

      <SectionCard title="ℹ️ Sistem Sinkronisasi Terintegrasi Buku Kas" subtitle="Berikut adalah alur otomatisasi aplikasi:">
        <div style={{ fontSize: 13, color: C.text, lineHeight: 1.7 }}>
          <p><strong>✓ Pencairan Jualan Manual Bersih:</strong> Ketika Anda menginput nominal bersih hasil klaim di tab 'Barang Keluar' dan mengklik 'Konfirmasi Cair', nominal riil manual tersebut langsung dilemparkan sebagai <strong>Debit (Uang Masuk)</strong> ke Buku Kas ini.</p>
          <p><strong>✓ Kasbon Karyawan → Kredit Otomatis:</strong> Sesaat setelah Anda mencatat nominal pinjaman kasbon pegawai di tab Keuangan, sistem langsung membuat baris <strong>Kredit (Uang Keluar)</strong> di Buku Kas ini demi menyeimbangkan neraca fisik keuangan toko.</p>
        </div>
      </SectionCard>
    </div>
  );
}

// ---------- SETTING MASTER KODE VARIABEL REGISTER ----------
const KATEGORI = [
  { key: "vendor", label: "Vendor Pabrik" },
  { key: "jenis", label: "Jenis Barang" },
  { key: "brand", label: "Brand / Merk" },
  { key: "bahan", label: "Bahan Kain" },
  { key: "warna", label: "Warna & Artikel" },
  { key: "ukuran", label: "Ukuran / Size" },
  { key: "karyawanMaster", label: "👥 Struktur Tim Pegawai" },
];

function MasterTab({ data, save, flash }) {
  const safeData = normalizeData(data || {});
  const [cat, setCat] = useState("vendor");
  const [form, setForm] = useState({ code: "", name: "" });
  const [formKaryawan, setFormKaryawan] = useState({ nama: "", peran: "Host Live" });

  const karyawanList = safeData.karyawanMaster || [];
  const listData = cat === "karyawanMaster" ? [] : (safeData.masterCode?.[cat] || []);

  const addKaryawan = () => {
    if (!formKaryawan.nama.trim()) { flash("Masukkan nama pegawai!"); return; }
    if (karyawanList.some((k) => k.nama?.toLowerCase() === formKaryawan.nama.trim().toLowerCase())) {
      flash("Nama pegawai tersebut sudah terdaftar!");
      return;
    }
    const newEmployee = {
      id: uid(),
      nama: formKaryawan.nama.toUpperCase().trim(),
      peran: formKaryawan.peran,
    };
    save({
      ...safeData,
      karyawanMaster: [...karyawanList, newEmployee],
    });
    setFormKaryawan({ nama: "", peran: "Host Live" });
    flash(`Personel ${newEmployee.nama} sukses dimasukkan ke struktur tim.`);
  };

  const removeKaryawan = (id) => {
    if (confirm("Hapus data pegawai ini dari hak sistem?")) {
      save({ ...safeData, karyawanMaster: karyawanList.filter((k) => k.id !== id) });
      flash("Pegawai dihapus.");
    }
  };

  const addMaster = () => {
    if (!form.code.trim() || !form.name.trim()) { flash("Isi inisial kode dan arti nama!"); return; }
    if (listData.some((x) => x.code?.toUpperCase() === form.code.trim().toUpperCase())) {
      flash("Inisial kode tersebut sudah terpakai di kategori ini!");
      return;
    }
    const next = {
      ...safeData,
      masterCode: {
        ...safeData.masterCode,
        [cat]: [...listData, { code: form.code.trim().toUpperCase(), name: form.name.trim().toUpperCase() }],
      },
    };
    save(next);
    setForm({ code: "", name: "" });
    flash("Kode master baru berhasil disimpan.");
  };

  const removeMaster = (code) => {
    if (confirm("Hapus parameter master code ini? Risiko SKU produk lama tidak terbaca jika dihapus.")) {
      save({
        ...safeData,
        masterCode: { ...safeData.masterCode, [cat]: listData.filter((x) => x.code !== code) },
      });
      flash("Parameter master dihapus.");
    }
  };

  return (
    <div>
      <SectionCard title="Setup Konfigurasi Kamus Kode Master" subtitle="Penyusunan parameter dasar akronim pembentuk kode SKU unik otomatis">
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 16, overflowX: "auto" }}>
          {KATEGORI.map((k) => (
            <button
              key={k.key}
              onClick={() => {
                setCat(k.key);
                setForm({ code: "", name: "" });
                setFormKaryawan({ nama: "", peran: "Host Live" });
              }}
              style={{
                padding: "7px 13px", borderRadius: 999, fontSize: 12.5, fontWeight: 600, cursor: "pointer",
                border: `1px solid ${cat === k.key ? C.primary : C.border}`,
                background: cat === k.key ? C.primary : "#fff", color: cat === k.key ? "#fff" : C.text, whiteSpace: "nowrap",
              }}
            >
              {k.label}
            </button>
          ))}
        </div>

        {/* SUB-SECTION RENDER GABUNGAN: KARYAWAN TIM */}
        {cat === "karyawanMaster" ? (
          <div>
            <div style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 13.5, fontWeight: 700, color: C.primaryDark, marginBottom: 12 }}>Daftarkan Anggota Staf Baru</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 10, alignItems: "end" }}>
                <Field label="Nama Lengkap Pegawai">
                  <input className="gm-input" value={formKaryawan.nama} onChange={(e) => setFormKaryawan({ ...formKaryawan, nama: e.target.value })} placeholder="Masukkan nama staf…" />
                </Field>
                <Field label="Tanggung Jawab Jabatan">
                  <select className="gm-select" value={formKaryawan.peran} onChange={(e) => setFormKaryawan({ ...formKaryawan, peran: e.target.value })}>
                    <option value="Host Live">Host Live Streaming</option>
                    <option value="Admin Live">Admin Live / Shift</option>
                    <option value="Owner">Direksi / Owner</option>
                  </select>
                </Field>
                <button className="gm-btn" onClick={addKaryawan}><Plus size={15} /> Registrasi Tim</button>
              </div>
            </div>
            <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr><th>Nama Pegawai Resmi</th><th>Tugas Divisi</th><th>Aksi</th></tr></thead>
              <tbody>
                {karyawanList.length === 0 && <EmptyRow colSpan={3} text="Belum ada anggota tim terdata." />}
                {karyawanList.map((k) => (
                  <tr key={k.id}>
                    <td style={{ fontWeight: 600 }}>{k.nama}</td>
                    <td><span style={{ fontSize: 11.5, fontWeight: 600, padding: "4px 9px", borderRadius: 6, background: k.peran === "Owner" ? "#FEE2E2" : "#DBEAFE", color: k.peran === "Owner" ? "#7F1D1D" : "#0C4A6E" }}>{k.peran}</span></td>
                    <td><DeleteBtn onClick={() => removeKaryawan(k.id)} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          /* SUB-SECTION PARAMETER SKEMA STANDARD */
          <div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr auto", gap: 10, marginBottom: 14 }}>
              <Field label="Singkatan Inisial (Kode SKU)">
                <input className="gm-input" value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} placeholder="cth: M" />
              </Field>
              <Field label="Jabaran Arti Nama Parameter">
                <input className="gm-input" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="cth: MATCHIATO" />
              </Field>
              <div style={{ display: "flex", alignItems: "end" }}>
                <button className="gm-btn" onClick={addMaster}><Plus size={15} /> Simpan Kode</button>
              </div>
            </div>
            <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr><th>Kode</th><th>Arti Parameter Kompilasi</th><th>Tindakan</th></tr></thead>
              <tbody>
                {listData.length === 0 && <EmptyRow colSpan={3} text="Belum ada data parameter terpasang." />}
                {listData.map((v) => (
                  <tr key={v.code}>
                    <td style={{ fontFamily: "monospace", fontWeight: 700 }}>{v.code}</td>
                    <td>{v.name}</td>
                    <td><DeleteBtn onClick={() => removeMaster(v.code)} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </SectionCard>
    </div>
  );
}
// ---------- RENDER LAPORAN DOKUMEN CETAK FISIK PDF ----------
function PrintTable({ title, rows }) {
  if (!rows || !rows.length) return null;
  const headers = Object.keys(rows[0]);
  return (
    <div style={{ marginBottom: 22, breakInside: "avoid" }}>
      <h3 style={{ fontFamily: "Georgia, serif", color: C.primaryDark, fontSize: 14, marginBottom: 6 }}>{title}</h3>
      <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "10px" }} border="1" cellPadding="5" borderColor="#ccc">
        <thead>
          <tr style={{ background: "#F3ECE1" }}>
            {headers.map((h) => <th key={h} style={{ textAlign: "left" }}>{h}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i}>
              {headers.map((h) => <td key={h}>{String(r[h] ?? "")}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function PrintReport({ data, stockMap, stokAkhir, onClose }) {
  const { produk, masuk, keluar, retur, reject, masterCode } = buildRows(data, stockMap, stokAkhir);
  
  useEffect(() => {
    const t = setTimeout(() => window.print(), 500);
    return () => clearTimeout(t);
  }, []);

  return (
    <div style={{ background: "#fff", color: "#1a1a1a", fontFamily: "sans-serif", padding: 30, maxWidth: 1100, margin: "0 auto" }}>
      <div className="no-print" style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginBottom: 20 }}>
        <button onClick={onClose} style={{ padding: "8px 14px", fontWeight: "bold", borderRadius: 8, cursor: "pointer" }}>← Kembali Ke Workspace</button>
        <button onClick={() => window.print()} style={{ padding: "8px 14px", fontWeight: "bold", border: "none", background: C.primary, color: "#fff", borderRadius: 8, cursor: "pointer" }}>Cetak / Simpan PDF</button>
      </div>
      <h2 style={{ fontFamily: "Georgia, serif", color: C.primaryDark, margin: 0 }}>Pembukuan Penjualan & Inventaris Gamis Niki Dziyab</h2>
      <small style={{ color: C.muted }}>Waktu Penarikan Berkas Laporan: {todayStr()}</small>
      <hr style={{ margin: "15px 0 25px" }} />
      
      <PrintTable title="1. Ringkasan Aset Saldo Unit Persediaan Gudang" rows={produk} />
      <PrintTable title="2. Dokumen Kelola Mutasi Barang Masuk" rows={masuk} />
      <PrintTable title="3. Laporan Buku Kas Keluar (Transaksi Penjualan)" rows={keluar} />
      <PrintTable title="4. Rekapitulasi Klaim Retur Masalah Pembeli" rows={retur} />
      <PrintTable title="5. Buku Catatan Afkir / Reject Produksi Kain" rows={reject} />
      <PrintTable title="6. Kamus Parameter Kode Master Aktif" rows={masterCode} />
      
      <style>{`
        @media print {
          .no-print { display: none !important; }
          @page { margin: 12mm; }
        }
      `}</style>
    </div>
  );
}
