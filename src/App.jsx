import { db } from "./firebase";
import { collection, onSnapshot, addDoc, deleteDoc, doc, setDoc } from "firebase/firestore";
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
  Upload,
  FileSpreadsheet,
  FileText,
  Printer,
  ChevronUp,
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

// ---------- design tokens ----------
const C = {
  bg: "#FAF6EF",
  surface: "#FFFFFF",
  surfaceAlt: "#F3ECE1",
  primary: "#7A2E3B",
  primaryDark: "#54202A",
  primarySoft: "#F1E1E1",
  accent: "#C69436",
  accentSoft: "#F3E4C2",
  text: "#2B1D1F",
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

const TOKO_OPTIONS = ["Shopee", "TikTok", "Tokopedia", "Lazada", "Offline"];
const PENGIRIMAN_OPTIONS = ["Shopee", "TikTokShop", "Tokopedia", "Lazada", "Manual"];
const ONGKIR_OPTIONS = ["Potong di Aplikasi", "Manual"];

const STORAGE_KEY = "gamis-inventory-v1";

const normalizeData = (value) => {
  const emptyMaster = { vendor: [], jenis: [], brand: [], bahan: [], warna: [], ukuran: [] };
  const source = value && typeof value === "object" ? value : {};
  const master = source.masterCode && typeof source.masterCode === "object" ? source.masterCode : {};

  return {
    products: Array.isArray(source.products) ? source.products : [],
    masuk: Array.isArray(source.masuk) ? source.masuk : [],
    keluar: Array.isArray(source.keluar) ? source.keluar : [],
    retur: Array.isArray(source.retur) ? source.retur : [],
    reject: Array.isArray(source.reject) ? source.reject : [],
    masterCode: {
      vendor: Array.isArray(master.vendor) ? master.vendor : emptyMaster.vendor,
      jenis: Array.isArray(master.jenis) ? master.jenis : emptyMaster.jenis,
      brand: Array.isArray(master.brand) ? master.brand : emptyMaster.brand,
      bahan: Array.isArray(master.bahan) ? master.bahan : emptyMaster.bahan,
      warna: Array.isArray(master.warna) ? master.warna : emptyMaster.warna,
      ukuran: Array.isArray(master.ukuran) ? master.ukuran : emptyMaster.ukuran,
    },
  };
};

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
  // Rem darurat jika data belum siap
  if (!data) return { produk: [], masuk: [], keluar: [], retur: [], reject: [], masterCode: [], pencairan: [] };

  const masuk = (data.masuk || []).map((m) => ({
    // isi di dalam m tetap aman
    ...m
  }));

  const keluar = (data.keluar || []).map((k) => ({
    Tanggal: k.tanggal, "Kode Pesanan": k.kodePesanan, Shift: k.shift, Host: k.host,
    "Kode Barang": k.kode, "Nama Barang": ((data.products || []).find((p) => p.kode === k.kode) || {}).nama,
    Qty: k.qty, Toko: k.toko, "COD/NON COD": k.cod,
    Affiliate: k.affiliate ? "Ya" : "Tidak", "Nama Affiliate": k.namaAffiliate || "",
    Pengiriman: k.pengiriman || "", Ongkir: k.ongkir || "", Catatan: k.catatan || "",
    "Sudah Cair": k.cair ? "Ya" : "Belum",
  }));

  const retur = (data.retur || []).map((r) => ({
    Tanggal: r.tanggal, "Kode Pesanan": r.kodePesanan, "Kode Barang": r.kode, "Kode Retur": r.kodeRetur,
    Qty: r.qty, "Sesuai/Tidak": r.sesuai, Ekspedisi: r.ekspedisi, Catatan: r.catatan,
  }));

  const reject = (data.reject || []).map((r) => ({
    Tanggal: r.tanggal, "Kode Barang": r.kode, "Nama Barang": ((data.products || []).find((p) => p.kode === r.kode) || {}).nama,
    Qty: r.qty, Toko: r.toko, Catatan: r.catatan,
  }));

  const masterCode = [];
  // Perbaikan Baris 126: Menggunakan pengaman Array/Objek agar tidak crash
  const targetMaster = data.masterCode || [];
  const entries = Array.isArray(targetMaster) ? targetMaster : Object.entries(targetMaster);
  
  entries.forEach(([kategori, list]) => {
    if (list && typeof list.forEach === 'function') {
      list.forEach((item) => {
        masterCode.push({ Kategori: kategori.toUpperCase(), Kode: item.code, Nama: item.name });
      });
    }
  });

  // Jika masterCode kosong murni, isi dengan 1 baris keterangan agar .length di bawah tidak crash
  if (masterCode.length === 0) {
    masterCode.push({ Kategori: "BELUM ADA DATA", Kode: "-", Nama: "Silakan input kode master baru" });
  }

  const totalKeluar = (data.keluar || []).reduce((s, k) => s + (Number(k.qty) || 0), 0);
  const totalCair = (data.keluar || []).filter((k) => k.cair).reduce((s, k) => s + (Number(k.qty) || 0), 0);
  const pencairan = [{ "Barang Keluar": totalKeluar, "Sudah Cair": totalCair, Selisih: totalKeluar - totalCair }];
  
  return { produk: data.products || [], masuk, keluar, retur, reject, masterCode, pencairan };
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
  const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
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

function seedData() {
  return {
    masterCode: {
      vendor: [{ code: "A", name: "HABIB HUSEIN" }],
      jenis: [{ code: "1", name: "GAMIS" }],
      brand: [{ code: "1", name: "NIKI DZIYAB" }],
      bahan: [{ code: "1", name: "SABRINA SILK" }],
      warna: [
        { code: "1", name: "HITAM" },
        { code: "2", name: "MAHOGANI" },
        { code: "3", name: "EMRALD" },
        { code: "4", name: "MATCHIATO" },
        { code: "5", name: "BURGUNDY" },
      ],
      ukuran: [
        { code: "1", name: "L" },
        { code: "2", name: "XL" },
        { code: "3", name: "XXL" },
        { code: "4", name: "ALL SIZE" },
      ],
    },
    products: [
      { kode: "A11111", nama: "NIKI DZIYAB GAMIS HITAM L HABIB HUSEIN", jenis: "GAMIS", brand: "NIKI DZIYAB", bahan: "SABRINA SILK", warna: "HITAM", ukuran: "L", vendor: "HABIB HUSEIN" },
      { kode: "A11112", nama: "NIKI DZIYAB GAMIS HITAM XL HABIB HUSEIN", jenis: "GAMIS", brand: "NIKI DZIYAB", bahan: "SABRINA SILK", warna: "HITAM", ukuran: "XL", vendor: "HABIB HUSEIN" },
    ],
    masuk: [
      { id: uid(), tanggal: "2026-02-08", kode: "A11111", qty: 100, vendor: "HABIB HUSEIN", catatan: "Stok awal" },
      { id: uid(), tanggal: "2026-02-08", kode: "A11112", qty: 100, vendor: "HABIB HUSEIN", catatan: "Stok awal" },
    ],
    keluar: [
      { id: uid(), tanggal: "2026-02-08", kodePesanan: "HSUIAOHDISUOAH", shift: "", host: "", kode: "A11111", qty: 5, toko: "Shopee", cod: "COD", affiliate: false, namaAffiliate: "", pengiriman: "", ongkir: "", catatan: "", cair: false },
      { id: uid(), tanggal: "2026-02-08", kodePesanan: "260711D62ED3RC", shift: "", host: "", kode: "A11112", qty: 5, toko: "Shopee", cod: "COD", affiliate: false, namaAffiliate: "", pengiriman: "", ongkir: "", catatan: "", cair: false },
    ],
    retur: [],
    reject: [],
  };
}

const TABS = [
  { id: "dashboard", label: "Dasbor", icon: LayoutDashboard },
  { id: "produk", label: "Produk & Stok", icon: Shirt },
  { id: "masuk", label: "Barang Masuk", icon: ArrowDownToLine },
  { id: "keluar", label: "Barang Keluar", icon: ArrowUpFromLine },
  { id: "retur", label: "Retur", icon: RotateCcw },
  { id: "reject", label: "Reject", icon: XCircle },
  { id: "pencairan", label: "Pencairan", icon: Wallet },
  { id: "master", label: "Kode Master", icon: Settings2 },
];

export default function GamisInventoryApp() {
    const [data, setData] = useState(normalizeData({
    products: [],
    masuk: [],
    keluar: [],
    retur: [],
    reject: [],
    masterCode: { vendor: [], jenis: [], brand: [], bahan: [], warna: [], ukuran: [] }
  }));

  const [ready, setReady] = useState(false);
  const [tab, setTab] = useState("dashboard");
  const [toast, setToast] = useState("");
  const [printView, setPrintView] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const fileInputRef = useRef(null);

      useEffect(() => {
    // Membaca data dari Firestore secara realtime dari dokumen "data_utama"
    const unsubscribe = onSnapshot(doc(db, "pembukuan_gamis", "data_utama"), (snapshot) => {
      if (snapshot.exists()) {
        const cloudData = snapshot.data();
        setData(normalizeData(cloudData.pembukuanData));
      } else {
        setData(normalizeData({
          products: [],
          masuk: [],
          keluar: [],
          retur: [],
          reject: [],
          masterCode: { vendor: [], jenis: [], brand: [], bahan: [], warna: [], ukuran: [] }
        }));
      }
    });

    return () => unsubscribe();
  }, []);

  const save = async (next) => {
    const normalized = normalizeData(next);
    setData(normalized);
    
    try {
      // 2. Kirim dan simpan langsung ke dokumen "data_utama" di Firebase
      const docRef = doc(db, "pembukuan_gamis", "data_utama");
      await setDoc(docRef, {
        pembukuanData: normalized,
        updatedAt: new Date()
      });
      
      console.log("Data berhasil sinkron otomatis ke Cloud Firebase!");
    } catch (e) {
      console.error("Gagal menyimpan data ke Firebase:", e);
    }
  };


  const flash = (msg) => {
    setToast(msg);
    setTimeout(() => setToast(""), 2200);
  };

  useEffect(() => {
    if (!printView) return;
    const t = setTimeout(() => window.print(), 200);
    const onAfter = () => setPrintView(false);
    window.addEventListener("afterprint", onAfter);
    return () => { clearTimeout(t); window.removeEventListener("afterprint", onAfter); };
  }, [printView]);

    const saveAsFile = async () => {
    try {
      // Menghapus data lama di Firebase terlebih dahulu agar tidak menumpuk ganda
      // (Atau langsung memperbarui dokumen utama jika Anda menggunakan sistem satu dokumen)
      await addDoc(collection(db, "pembukuan_gamis"), {
        pembukuanData: data,
        updatedAt: new Date()
      });
      
      flash("Data berhasil disimpan ke Cloud!");
    } catch (error) {
      console.error("Gagal menyimpan data:", error);
      flash("Gagal menyimpan data ke internet!");
    }
  };


  const handleImportFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        if (!parsed.products || !parsed.masterCode) throw new Error("format tidak dikenali");
        save(parsed);
        flash("Data berhasil dimuat dari file");
      } catch (err) {
        flash("Gagal membaca file — pastikan file cadangan yang benar");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  const stockMap = useMemo(() => {
    if (!data) return {};
    const map = {};
    data.products.forEach((p) => {
      map[p.kode] = { masuk: 0, keluar: 0, returOk: 0, reject: 0 };
    });
    data.masuk.forEach((m) => { if (map[m.kode]) map[m.kode].masuk += Number(m.qty) || 0; });
    data.keluar.forEach((k) => { if (map[k.kode]) map[k.kode].keluar += Number(k.qty) || 0; });
    data.retur.forEach((r) => { if (map[r.kode] && r.sesuai === "Sesuai") map[r.kode].returOk += Number(r.qty) || 0; });
    data.reject.forEach((rj) => { if (map[rj.kode]) map[rj.kode].reject += Number(rj.qty) || 0; });
    return map;
  }, [data]);

  const stokAkhir = (kode) => {
    const s = stockMap[kode];
    if (!s) return 0;
    return s.masuk - s.keluar - s.reject + s.returOk;
  };

//  if (!ready || !data) {
//    return (
//      <div style={{ background: C.bg, minHeight: 480, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "IBM Plex Sans, sans-serif", color: C.muted }}>
//        Memuat data…
//      </div>
//    );
//  }

  if (printView) {
    return <PrintReport data={data} stockMap={stockMap} stokAkhir={stokAkhir} onClose={() => setPrintView(false)} />;
  }

  return (
    <div style={{ background: C.bg, minHeight: 640, fontFamily: "'IBM Plex Sans', sans-serif", color: C.text }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');
        .gm-h1 { font-family: 'Fraunces', serif; }
        .gm-input, .gm-select {
          background: #fff; border: 1px solid ${C.border}; border-radius: 8px;
          padding: 8px 10px; font-family: 'IBM Plex Sans', sans-serif; font-size: 13px;
          color: ${C.text}; width: 100%; outline: none; transition: border-color .15s;
        }
        .gm-input:focus, .gm-select:focus { border-color: ${C.primary}; }
        .gm-btn {
          background: ${C.primary}; color: #fff; border: none; border-radius: 8px;
          padding: 9px 16px; font-size: 13px; font-weight: 600; cursor: pointer;
          display: inline-flex; align-items: center; gap: 6px; transition: background .15s;
          font-family: 'IBM Plex Sans', sans-serif;
        }
        .gm-btn:hover { background: ${C.primaryDark}; }
        .gm-btn-ghost {
          background: transparent; color: ${C.primary}; border: 1px solid ${C.primary};
        }
        .gm-btn-ghost:hover { background: ${C.primarySoft}; }
        .gm-table th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: .04em; color: ${C.muted}; font-weight: 600; padding: 10px 12px; border-bottom: 1px solid ${C.border}; white-space: nowrap; }
        .gm-table td { padding: 10px 12px; border-bottom: 1px solid ${C.border}; font-size: 13px; vertical-align: middle; white-space: nowrap; }
        .gm-table tr:hover td { background: ${C.surfaceAlt}; }
        .gm-tab-scroll::-webkit-scrollbar { height: 4px; }
        .gm-tab-scroll::-webkit-scrollbar-thumb { background: ${C.border}; border-radius: 4px; }
        .gm-card { background: ${C.surface}; border: 1px solid ${C.border}; border-radius: 14px; }
      `}</style>

      {menuOpen && <div onClick={() => setMenuOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 20 }} />}

      {/* header */}
      <div style={{ borderBottom: `1px solid ${C.border}`, background: C.surface, position: "sticky", top: 0, zIndex: 10 }}>
        <div style={{ padding: "18px 24px 0" }}>
          <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{ width: 34, height: 34, borderRadius: 9, background: C.primary, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <Shirt size={18} color={C.accentSoft} />
              </div>
              <div>
                <div className="gm-h1" style={{ fontSize: 19, fontWeight: 700, lineHeight: 1.1, color: C.primaryDark }}>Pembukuan Gamis Niki Dziyab</div>
                <div style={{ fontSize: 11.5, color: C.muted }}>Stok · Barang Masuk & Keluar · Retur · Reject · Pencairan</div>
              </div>
            </div>

            <div style={{ display: "flex", alignItems: "center", gap: 6, position: "relative" }}>
              <input ref={fileInputRef} type="file" accept="application/json" onChange={handleImportFile} style={{ display: "none" }} />
              <button className="gm-btn-ghost" style={{ borderRadius: 8, padding: "8px 12px", fontSize: 12.5, fontWeight: 600, display: "inline-flex", alignItems: "center", gap: 6, cursor: "pointer", background: "transparent" }} onClick={() => fileInputRef.current?.click()} title="Muat file cadangan (.json)">
                <Upload size={14} /> Impor
              </button>
              <button className="gm-btn" style={{ padding: "8px 14px", fontSize: 12.5 }} onClick={saveAsFile} title="Unduh cadangan data sebagai file .json">
                <Save size={14} /> Simpan File
              </button>
              <div style={{ position: "relative" }}>
                <button
                  className="gm-btn"
                  style={{ padding: "8px 14px", fontSize: 12.5, background: C.accent }}
                  onClick={() => setMenuOpen((v) => !v)}
                >
                  Export {menuOpen ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
                </button>
                {menuOpen && (
                  <div style={{ position: "absolute", right: 0, top: "calc(100% + 6px)", background: "#fff", border: `1px solid ${C.border}`, borderRadius: 10, boxShadow: "0 10px 28px rgba(0,0,0,.14)", overflow: "hidden", zIndex: 30, minWidth: 190 }}>
                    <MenuItem icon={FileSpreadsheet} label="Export ke Excel (.xlsx)" onClick={() => { exportExcel(data, stockMap, stokAkhir); setMenuOpen(false); flash("File Excel diunduh"); }} />
                    <MenuItem icon={FileText} label="Export ke Word (.doc)" onClick={() => { exportWord(data, stockMap, stokAkhir); setMenuOpen(false); flash("File Word diunduh"); }} />
                    <MenuItem icon={Printer} label="Export ke PDF" onClick={() => { setMenuOpen(false); setPrintView(true); }} />
                  </div>
                )}
              </div>
            </div>
          </div>
          <div className="gm-tab-scroll" style={{ display: "flex", gap: 4, marginTop: 16, overflowX: "auto", paddingBottom: 0 }}>
            {TABS.map((t) => {
              const Icon = t.icon;
              const active = tab === t.id;
              return (
                <button
                  key={t.id}
                  onClick={() => setTab(t.id)}
                  style={{
                    display: "flex", alignItems: "center", gap: 6, whiteSpace: "nowrap",
                    padding: "9px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer",
                    background: "transparent", border: "none",
                    color: active ? C.primary : C.muted,
                    borderBottom: active ? `2.5px solid ${C.primary}` : "2.5px solid transparent",
                    fontFamily: "'IBM Plex Sans', sans-serif",
                  }}
                >
                  <Icon size={15} />
                  {t.label}
                </button>
              );
            })}
          </div>
        </div>
      </div>

      <div style={{ padding: "22px 24px 60px", maxWidth: 1180, margin: "0 auto" }}>
        {tab === "dashboard" && <Dashboard data={data} stockMap={stockMap} stokAkhir={stokAkhir} />}
        {tab === "produk" && <ProdukTab data={data} save={save} stockMap={stockMap} stokAkhir={stokAkhir} flash={flash} />}
        {tab === "masuk" && <MasukTab data={data} save={save} flash={flash} />}
        {tab === "keluar" && <KeluarTab data={data} save={save} stokAkhir={stokAkhir} flash={flash} />}
        {tab === "retur" && <ReturTab data={data} save={save} flash={flash} />}
        {tab === "reject" && <RejectTab data={data} save={save} flash={flash} />}
        {tab === "pencairan" && <PencairanTab data={data} save={save} flash={flash} />}
        {tab === "master" && <MasterTab data={data} save={save} flash={flash} />}
      </div>

      {toast && (
        <div style={{ position: "fixed", bottom: 20, left: "50%", transform: "translateX(-50%)", background: C.primaryDark, color: "#fff", padding: "10px 18px", borderRadius: 999, fontSize: 13, fontWeight: 500, boxShadow: "0 8px 24px rgba(0,0,0,.2)" }}>
          {toast}
        </div>
      )}
    </div>
  );
}

// ---------- shared bits ----------
function Field({ label, children }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
      <label style={{ fontSize: 11.5, fontWeight: 600, color: C.muted, textTransform: "uppercase", letterSpacing: ".03em" }}>{label}</label>
      {children}
    </div>
  );
}

function SectionCard({ title, subtitle, right, children }) {
  return (
    <div className="gm-card" style={{ padding: 18, marginBottom: 18 }}>
      {(title || right) && (
        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 14, gap: 12, flexWrap: "wrap" }}>
          <div>
            {title && <div className="gm-h1" style={{ fontSize: 16.5, fontWeight: 600, color: C.primaryDark }}>{title}</div>}
            {subtitle && <div style={{ fontSize: 12.5, color: C.muted, marginTop: 2 }}>{subtitle}</div>}
          </div>
          {right}
        </div>
      )}
      {children}
    </div>
  );
}

function ProductSelect({ products, value, onChange }) {
  return (
    <select className="gm-select" value={value} onChange={(e) => onChange(e.target.value)}>
      <option value="">Pilih produk…</option>
      {products.map((p) => (
        <option key={p.kode} value={p.kode}>{p.kode} — {p.nama}</option>
      ))}
    </select>
  );
}

function EmptyRow({ colSpan, text }) {
  return (
    <tr>
      <td colSpan={colSpan} style={{ padding: "26px 12px", textAlign: "center", color: C.muted, fontSize: 13 }}>{text}</td>
    </tr>
  );
}

function MenuItem({ icon: Icon, label, onClick }) {
  return (
    <button
      onClick={onClick}
      style={{
        display: "flex", alignItems: "center", gap: 9, width: "100%", padding: "10px 14px",
        fontSize: 12.5, fontWeight: 500, background: "transparent", border: "none", cursor: "pointer",
        color: C.text, textAlign: "left", fontFamily: "'IBM Plex Sans', sans-serif",
      }}
      onMouseEnter={(e) => (e.currentTarget.style.background = C.surfaceAlt)}
      onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
    >
      <Icon size={15} color={C.primary} /> {label}
    </button>
  );
}

function DeleteBtn({ onClick }) {
  return (
    <button onClick={onClick} title="Hapus" style={{ background: "transparent", border: "none", cursor: "pointer", color: C.muted, padding: 4 }}>
      <Trash2 size={14} />
    </button>
  );
}

// ---------- Dashboard ----------
function Dashboard({ data, stockMap, stokAkhir }) {
  const totalProduk = data.products.length;
  const totalStok = data.products.reduce((s, p) => s + stokAkhir(p.kode), 0);
  const bulanIni = monthOf(todayStr());
  const keluarBulanIni = data.keluar.filter((k) => monthOf(k.tanggal) === bulanIni).reduce((s, k) => s + (Number(k.qty) || 0), 0);
  const returBulanIni = data.retur.filter((r) => monthOf(r.tanggal) === bulanIni).reduce((s, r) => s + (Number(r.qty) || 0), 0);
  const rejectBulanIni = data.reject.filter((r) => monthOf(r.tanggal) === bulanIni).reduce((s, r) => s + (Number(r.qty) || 0), 0);

  const totalKeluar = data.keluar.reduce((s, k) => s + (Number(k.qty) || 0), 0);
  const totalCair = data.keluar.filter((k) => k.cair).reduce((s, k) => s + (Number(k.qty) || 0), 0);

  const chartData = BULAN_ID.slice(6, 12).map((b) => ({
    bulan: b.slice(0, 3),
    qty: data.keluar.filter((k) => monthOf(k.tanggal) === b).reduce((s, k) => s + (Number(k.qty) || 0), 0),
  }));

  const lowStock = data.products.filter((p) => stokAkhir(p.code) <= 10);

  const cards = [
    { label: "Total Produk", value: fmt(totalProduk), color: C.primary },
    { label: "Stok Tersedia", value: fmt(totalStok), color: C.success },
    { label: `Keluar — ${bulanIni}`, value: fmt(keluarBulanIni), color: C.accent },
    { label: `Retur — ${bulanIni}`, value: fmt(returBulanIni), color: C.danger },
    { label: `Reject — ${bulanIni}`, value: fmt(rejectBulanIni), color: C.muted },
  ];

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 18 }}>
        {cards.map((c) => (
          <div key={c.label} className="gm-card" style={{ padding: 16 }}>
            <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600, textTransform: "uppercase", letterSpacing: ".03em" }}>{c.label}</div>
            <div className="gm-h1" style={{ fontSize: 26, fontWeight: 700, color: c.color, marginTop: 6 }}>{c.value}</div>
          </div>
        ))}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1.5fr 1fr", gap: 16, alignItems: "stretch" }}>
        <SectionCard title="Barang Keluar per Bulan" subtitle="Agustus – Desember">
          <div style={{ height: 220 }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartData} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={C.border} vertical={false} />
                <XAxis dataKey="bulan" tick={{ fontSize: 12, fill: C.muted }} axisLine={{ stroke: C.border }} tickLine={false} />
                <YAxis tick={{ fontSize: 12, fill: C.muted }} axisLine={false} tickLine={false} />
                <Tooltip contentStyle={{ borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 12 }} />
                <Bar dataKey="qty" radius={[5, 5, 0, 0]}>
                  {chartData.map((entry, i) => (
                    <Cell key={i} fill={entry.bulan === bulanIni.slice(0, 3) ? C.primary : C.accentSoft} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </SectionCard>

        <SectionCard title="Pencairan" subtitle="Barang keluar vs sudah cair">
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <Stat label="Barang Keluar" value={fmt(totalKeluar)} />
            <Stat label="Sudah Cair" value={fmt(totalCair)} color={C.success} />
            <Stat label="Selisih" value={fmt(totalKeluar - totalCair)} color={C.danger} />
          </div>
        </SectionCard>
      </div>

      <SectionCard title="Stok Menipis" subtitle="Produk dengan stok akhir ≤ 10" right={lowStock.length > 0 && <AlertTriangle size={16} color={C.danger} />}>
        {lowStock.length === 0 ? (
          <div style={{ fontSize: 13, color: C.muted }}>Tidak ada produk dengan stok menipis.</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {lowStock.map((p) => (
              <div key={p.kode} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 10px", background: C.surfaceAlt, borderRadius: 8 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ width: 10, height: 10, borderRadius: "50%", background: WARNA_SWATCH[p.warna] || C.muted, display: "inline-block" }} />
                  <span style={{ fontSize: 13, fontWeight: 500 }}>{p.kode} — {p.nama}</span>
                </div>
                <span style={{ fontSize: 13, fontWeight: 700, color: C.danger }}>{stokAkhir(p.kode)}</span>
              </div>
            ))}
          </div>
        )}
      </SectionCard>
    </div>
  );
}

function Stat({ label, value, color }) {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
      <span style={{ fontSize: 13, color: C.muted }}>{label}</span>
      <span className="gm-h1" style={{ fontSize: 17, fontWeight: 700, color: color || C.text }}>{value}</span>
    </div>
  );
}

// ---------- Produk & Stok ----------
function ProdukTab({ data, save, stockMap, stokAkhir, flash }) {
  const [q, setQ] = useState("");
  const safeData = data || { products: [], masuk: [], keluar: [], retur: [], reject: [], masterCode: { vendor: [], jenis: [], brand: [], bahan: [], warna: [], ukuran: [] } };
  const products = Array.isArray(safeData.products) ? safeData.products : [];
  const mc = safeData.masterCode || { vendor: [], jenis: [], brand: [], bahan: [], warna: [], ukuran: [] };
  const [form, setForm] = useState({ vendor: "", jenis: "", brand: "", bahan: "", warna: "", ukuran: "", nama: "", stokAwal: "" });

  const codeOf = (list, name) => (Array.isArray(list) ? (list.find((x) => x.name === name) || {}).code : "") || "";
  const previewKode =
    form.vendor && form.jenis && form.brand && form.bahan && form.warna && form.ukuran
      ? codeOf(mc.vendor, form.vendor) + codeOf(mc.jenis, form.jenis) + codeOf(mc.brand, form.brand) + codeOf(mc.bahan, form.bahan) + codeOf(mc.warna, form.warna) + codeOf(mc.ukuran, form.ukuran)
      : "";

  const addProduk = () => {
    if (!previewKode) { flash("Lengkapi vendor, jenis, brand, bahan, warna, dan ukuran dulu"); return; }
    if (products.some((p) => p.kode === previewKode)) { flash("Kode barang sudah ada"); return; }
    const nama = form.nama.trim() || `${form.brand} ${form.jenis} ${form.warna} ${form.ukuran} ${form.vendor}`.toUpperCase();
    const next = { ...safeData, products: [...products, { kode: previewKode, nama, jenis: form.jenis, brand: form.brand, bahan: form.bahan, warna: form.warna, ukuran: form.ukuran, vendor: form.vendor }] };
    if (Number(form.stokAwal) > 0) {
      next.masuk = [...(safeData.masuk || []), { id: uid(), tanggal: todayStr(), kode: previewKode, qty: Number(form.stokAwal), vendor: form.vendor, catatan: "Stok awal" }];
    }
    save(next);
    setForm({ vendor: "", jenis: "", brand: "", bahan: "", warna: "", ukuran: "", nama: "", stokAwal: "" });
    flash(`Produk ${previewKode} ditambahkan`);
  };

  const removeProduk = (kode) => {
    const hasLogs = [...(safeData.masuk || []), ...(safeData.keluar || []), ...(safeData.retur || []), ...(safeData.reject || [])].some((r) => r && r.kode === kode);
    if (hasLogs) { flash("Tidak bisa dihapus — produk ini punya riwayat transaksi"); return; }
    save({ ...safeData, products: products.filter((p) => p.kode !== kode) });
  };

  const filtered = products.filter((p) => {
    if (!p) return false;
    const kataKunci = q.trim().toLowerCase();
    if (!kataKunci) return true;
    const kode = String(p.kode || "").toLowerCase();
    const nama = String(p.nama || "").toLowerCase();
    const warna = String(p.warna || "").toLowerCase();
    const brand = String(p.brand || "").toLowerCase();
    return kode.includes(kataKunci) || nama.includes(kataKunci) || warna.includes(kataKunci) || brand.includes(kataKunci    
    );
  });

  return (
    <div>
      <SectionCard title="Tambah Produk" subtitle="Kode dibuat otomatis dari Kode Master (Vendor + Jenis + Brand + Bahan + Warna + Ukuran)">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px,1fr))", gap: 10 }}>
          <Field label="Vendor">
            <select className="gm-select" value={form.vendor} onChange={(e) => setForm({ ...form, vendor: e.target.value })}>
              <option value="">Pilih…</option>
              {(mc.vendor || []).map((v) => <option key={v.code} value={v.name}>{v.name}</option>)}
            </select>
          </Field>
          <Field label="Jenis Barang">
            <select className="gm-select" value={form.jenis} onChange={(e) => setForm({ ...form, jenis: e.target.value })}>
              <option value="">Pilih…</option>
              {(mc.jenis || []).map((v) => <option key={v.code} value={v.name}>{v.name}</option>)}
            </select>
          </Field>
          <Field label="Brand">
            <select className="gm-select" value={form.brand} onChange={(e) => setForm({ ...form, brand: e.target.value })}>
              <option value="">Pilih…</option>
              {(mc.brand || []).map((v) => <option key={v.code} value={v.name}>{v.name}</option>)}
            </select>
          </Field>
          <Field label="Bahan">
            <select className="gm-select" value={form.bahan} onChange={(e) => setForm({ ...form, bahan: e.target.value })}>
              <option value="">Pilih…</option>
              {(mc.bahan || []).map((v) => <option key={v.code} value={v.name}>{v.name}</option>)}
            </select>
          </Field>
          <Field label="Warna & Artikel">
            <select className="gm-select" value={form.warna} onChange={(e) => setForm({ ...form, warna: e.target.value })}>
              <option value="">Pilih…</option>
              {(mc.warna || []).map((v) => <option key={v.code} value={v.name}>{v.name}</option>)}
            </select>
          </Field>
          <Field label="Ukuran">
            <select className="gm-select" value={form.ukuran} onChange={(e) => setForm({ ...form, ukuran: e.target.value })}>
              <option value="">Pilih…</option>
              {(mc.ukuran || []).map((v) => <option key={v.code} value={v.name}>{v.name}</option>)}
            </select>
          </Field>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr auto", gap: 10, marginTop: 10, alignItems: "end" }}>
          <Field label="Nama Barang (opsional, otomatis jika kosong)">
            <input className="gm-input" value={form.nama} onChange={(e) => setForm({ ...form, nama: e.target.value })} placeholder="Nama barang lengkap" />
          </Field>
          <Field label="Stok Awal">
            <input className="gm-input" type="number" min="0" value={form.stokAwal} onChange={(e) => setForm({ ...form, stokAwal: e.target.value })} placeholder="0" />
          </Field>
          <button className="gm-btn" onClick={addProduk}><Plus size={15} /> Tambah</button>
        </div>
        {previewKode && (
          <div style={{ marginTop: 10, fontSize: 12.5, color: C.muted }}>
            Kode barang: <span style={{ fontWeight: 700, color: C.primary, fontFamily: "monospace", fontSize: 13.5 }}>{previewKode}</span>
          </div>
        )}
      </SectionCard>

      <SectionCard
        title="Daftar Produk & Stok"
        right={
          <div style={{ position: "relative", width: 220 }}>
            <Search size={14} style={{ position: "absolute", left: 9, top: 10, color: C.muted }} />
            <input className="gm-input" style={{ paddingLeft: 28 }} placeholder="Cari kode / nama / warna" value={q} onChange={(e) => setQ(e.target.value)} />
          </div>
        }
      >
        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th>Kode</th><th>Nama Barang</th><th>Warna</th><th>Ukuran</th><th>Vendor</th>
                <th>Masuk</th><th>Terjual</th><th>Retur OK</th><th>Reject</th><th>Stok Akhir</th><th></th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 && <EmptyRow colSpan={11} text="Belum ada produk." />}
              {filtered.map((p) => {
                const s = stockMap[p.kode] || { masuk: 0, keluar: 0, returOk: 0, reject: 0 };
                const akhir = stokAkhir(p.kode);
                return (
                  <tr key={p.kode}>
                    <td style={{ fontFamily: "monospace", fontWeight: 700, color: C.primary }}>{p.kode}</td>
                    <td style={{ whiteSpace: "normal", minWidth: 220 }}>{p.nama}</td>
                    <td><span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}><span style={{ width: 9, height: 9, borderRadius: "50%", background: WARNA_SWATCH[p.warna] || C.muted, display: "inline-block" }} />{p.warna}</span></td>
                    <td>{p.ukuran}</td>
                    <td>{p.vendor}</td>
                    <td>{fmt(s.masuk)}</td>
                    <td>{fmt(s.keluar)}</td>
                    <td>{fmt(s.returOk)}</td>
                    <td>{fmt(s.reject)}</td>
                    <td style={{ fontWeight: 700, color: akhir <= 10 ? C.danger : C.success }}>{fmt(akhir)}</td>
                    <td><DeleteBtn onClick={() => removeProduk(p.kode)} /></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </SectionCard>
    </div>
  );
}

// ---------- Barang Masuk ----------
function MasukTab({ data, save, flash }) {
  const [form, setForm] = useState({ tanggal: todayStr(), kode: "", qty: "", catatan: "" });

  const submit = () => {
    if (!form.kode || !Number(form.qty)) { flash("Pilih produk dan isi qty"); return; }
    const prod = data.products.find((p) => p.kode === form.kode);
    const entry = { id: uid(), tanggal: form.tanggal, kode: form.kode, qty: Number(form.qty), vendor: prod?.vendor || "", catatan: form.catatan };
    save({ ...data, masuk: [entry, ...data.masuk] });
    setForm({ tanggal: form.tanggal, kode: "", qty: "", catatan: "" });
    flash("Barang masuk dicatat");
  };

  const remove = (id) => save({ ...data, masuk: data.masuk.filter((m) => m.id !== id) });

  return (
    <div>
      <SectionCard title="Catat Barang Masuk">
        <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr 1fr 1.5fr auto", gap: 10, alignItems: "end" }}>
          <Field label="Tanggal"><input className="gm-input" type="date" value={form.tanggal} onChange={(e) => setForm({ ...form, tanggal: e.target.value })} /></Field>
          <Field label="Produk"><ProductSelect products={data.products} value={form.kode} onChange={(v) => setForm({ ...form, kode: v })} /></Field>
          <Field label="Qty"><input className="gm-input" type="number" min="0" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} /></Field>
          <Field label="Catatan"><input className="gm-input" value={form.catatan} onChange={(e) => setForm({ ...form, catatan: e.target.value })} placeholder="Opsional" /></Field>
          <button className="gm-btn" onClick={submit}><Plus size={15} /> Simpan</button>
        </div>
      </SectionCard>

      <SectionCard title="Riwayat Barang Masuk">
        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th>Tanggal</th><th>Kode</th><th>Nama Barang</th><th>Qty</th><th>Vendor</th><th>Catatan</th><th></th></tr></thead>
            <tbody>
              {data.masuk.length === 0 && <EmptyRow colSpan={7} text="Belum ada catatan barang masuk." />}
              {data.masuk.map((m) => {
                const prod = data.products.find((p) => p.kode === m.kode);
                return (
                  <tr key={m.id}>
                    <td>{m.tanggal}</td>
                    <td style={{ fontFamily: "monospace", fontWeight: 700, color: C.primary }}>{m.kode}</td>
                    <td style={{ whiteSpace: "normal", minWidth: 200 }}>{prod?.nama || "—"}</td>
                    <td>{fmt(m.qty)}</td>
                    <td>{m.vendor}</td>
                    <td>{m.catatan}</td>
                    <td><DeleteBtn onClick={() => remove(m.id)} /></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </SectionCard>
    </div>
  );
}

// ---------- Barang Keluar ----------
function KeluarTab({ data, save, stokAkhir, flash }) {
  const emptyForm = { tanggal: todayStr(), kodePesanan: "", shift: "", host: "", kode: "", qty: "", toko: "Shopee", cod: "COD", affiliate: false, namaAffiliate: "", pengiriman: "", ongkir: "", catatan: "" };
  const [form, setForm] = useState(emptyForm);
  const [monthFilter, setMonthFilter] = useState("SEMUA");

  const submit = () => {
    if (!form.kode || !Number(form.qty)) { flash("Pilih produk dan isi qty"); return; }
    const sisa = stokAkhir(form.kode);
    if (Number(form.qty) > sisa) { flash(`Stok tidak cukup (tersisa ${sisa})`); return; }
    if (form.affiliate && (!form.pengiriman || !form.ongkir)) { flash("Lengkapi pengiriman & ongkir untuk pesanan affiliate"); return; }
    const entry = { id: uid(), ...form, qty: Number(form.qty), cair: false };
    save({ ...data, keluar: [entry, ...data.keluar] });
    setForm({ ...emptyForm, tanggal: form.tanggal });
    flash("Barang keluar dicatat");
  };

  const remove = (id) => save({ ...data, keluar: data.keluar.filter((k) => k.id !== id) });

  const months = ["SEMUA", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"];
  const rows = data.keluar.filter((k) => monthFilter === "SEMUA" || monthOf(k.tanggal) === monthFilter);
  const totalQty = rows.reduce((s, k) => s + (Number(k.qty) || 0), 0);

  return (
    <div>
      <SectionCard title="Catat Barang Keluar" subtitle="Stok akan otomatis berkurang">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px,1fr))", gap: 10 }}>
          <Field label="Tanggal"><input className="gm-input" type="date" value={form.tanggal} onChange={(e) => setForm({ ...form, tanggal: e.target.value })} /></Field>
          <Field label="Kode Pesanan"><input className="gm-input" value={form.kodePesanan} onChange={(e) => setForm({ ...form, kodePesanan: e.target.value })} /></Field>
          <Field label="Shift"><input className="gm-input" value={form.shift} onChange={(e) => setForm({ ...form, shift: e.target.value })} placeholder="Pagi / Siang / Malam" /></Field>
          <Field label="Host"><input className="gm-input" value={form.host} onChange={(e) => setForm({ ...form, host: e.target.value })} /></Field>
          <Field label="Toko">
            <select className="gm-select" value={form.toko} onChange={(e) => setForm({ ...form, toko: e.target.value })}>
              {TOKO_OPTIONS.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
          </Field>
          <Field label="COD / Non COD">
            <select className="gm-select" value={form.cod} onChange={(e) => setForm({ ...form, cod: e.target.value })}>
              <option>COD</option><option>NON COD</option>
            </select>
          </Field>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr auto", gap: 10, marginTop: 10, alignItems: "end" }}>
          <Field label="Produk"><ProductSelect products={data.products} value={form.kode} onChange={(v) => setForm({ ...form, kode: v })} /></Field>
          <Field label="Qty"><input className="gm-input" type="number" min="0" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} /></Field>
          <button className="gm-btn" onClick={submit}><Plus size={15} /> Simpan</button>
        </div>

        <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 16, fontSize: 13, fontWeight: 600, color: C.primaryDark, cursor: "pointer" }}>
          <input type="checkbox" checked={form.affiliate} onChange={(e) => setForm({ ...form, affiliate: e.target.checked })} style={{ width: 15, height: 15, accentColor: C.primary }} />
          Pesanan Affiliate
        </label>

        {form.affiliate && (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px,1fr))", gap: 10, marginTop: 10, padding: 12, background: C.surfaceAlt, borderRadius: 10 }}>
            <Field label="Nama Affiliate"><input className="gm-input" value={form.namaAffiliate} onChange={(e) => setForm({ ...form, namaAffiliate: e.target.value })} placeholder="Nama / akun affiliate" /></Field>
            <Field label="Pengiriman melalui">
              <select className="gm-select" value={form.pengiriman} onChange={(e) => setForm({ ...form, pengiriman: e.target.value })}>
                <option value="">Pilih…</option>
                {PENGIRIMAN_OPTIONS.map((p) => <option key={p} value={p}>{p === "Manual" ? "Pengiriman Manual" : `Aplikasi ${p}`}</option>)}
              </select>
            </Field>
            <Field label="Ongkir">
              <select className="gm-select" value={form.ongkir} onChange={(e) => setForm({ ...form, ongkir: e.target.value })}>
                <option value="">Pilih…</option>
                {ONGKIR_OPTIONS.map((o) => <option key={o} value={o}>{o}</option>)}
              </select>
            </Field>
          </div>
        )}

        <div style={{ marginTop: 10 }}>
          <Field label="Catatan"><input className="gm-input" value={form.catatan} onChange={(e) => setForm({ ...form, catatan: e.target.value })} placeholder="Opsional" /></Field>
        </div>
      </SectionCard>

      <SectionCard
        title="Riwayat Barang Keluar"
        subtitle={`Total qty: ${fmt(totalQty)}`}
        right={
          <select className="gm-select" style={{ width: 160 }} value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)}>
            {months.map((m) => <option key={m} value={m}>{m}</option>)}
          </select>
        }
      >
        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th>Tanggal</th><th>Kode Pesanan</th><th>Kode</th><th>Produk</th><th>Qty</th><th>Toko</th><th>COD</th><th>Affiliate</th><th>Catatan</th><th>Cair</th><th></th></tr></thead>
            <tbody>
              {rows.length === 0 && <EmptyRow colSpan={11} text="Belum ada catatan barang keluar." />}
              {rows.map((k) => {
                const prod = data.products.find((p) => p.kode === k.kode);
                return (
                  <tr key={k.id}>
                    <td>{k.tanggal}</td>
                    <td>{k.kodePesanan}</td>
                    <td style={{ fontFamily: "monospace", fontWeight: 700, color: C.primary }}>{k.kode}</td>
                    <td style={{ whiteSpace: "normal", minWidth: 180 }}>{prod?.nama || "—"}</td>
                    <td>{fmt(k.qty)}</td>
                    <td>{k.toko}</td>
                    <td>{k.cod}</td>
                    <td style={{ whiteSpace: "normal", minWidth: 190 }}>
                      {k.affiliate ? (
                        <span style={{ fontSize: 11.5 }}>
                          <span style={{ fontWeight: 700, color: C.primary }}>{k.namaAffiliate || "Affiliate"}</span>
                          {k.pengiriman && <> · {k.pengiriman === "Manual" ? "Pengiriman Manual" : `Aplikasi ${k.pengiriman}`}</>}
                          {k.ongkir && <> · Ongkir: {k.ongkir}</>}
                        </span>
                      ) : "—"}
                    </td>
                    <td style={{ whiteSpace: "normal", minWidth: 120 }}>{k.catatan}</td>
                    <td>
                      <button
                        onClick={() => save({ ...data, keluar: data.keluar.map((x) => x.id === k.id ? { ...x, cair: !x.cair } : x) })}
                        style={{ fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 999, border: "none", cursor: "pointer", background: k.cair ? C.success : C.surfaceAlt, color: k.cair ? "#fff" : C.muted }}
                      >
                        {k.cair ? "Sudah" : "Belum"}
                      </button>
                    </td>
                    <td><DeleteBtn onClick={() => remove(k.id)} /></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </SectionCard>
    </div>
  );
}

// ---------- Retur ----------
function ReturTab({ data, save, flash }) {
  const [form, setForm] = useState({ tanggal: todayStr(), kodePesanan: "", kode: "", kodeRetur: "", qty: "", sesuai: "Sesuai", ekspedisi: "", catatan: "" });

  const submit = () => {
    if (!form.kode || !Number(form.qty)) { flash("Pilih produk dan isi qty"); return; }
    save({ ...data, retur: [{ id: uid(), ...form, qty: Number(form.qty) }, ...data.retur] });
    setForm({ tanggal: form.tanggal, kodePesanan: "", kode: "", kodeRetur: "", qty: "", sesuai: "Sesuai", ekspedisi: "", catatan: "" });
    flash("Retur dicatat");
  };
  const remove = (id) => save({ ...data, retur: data.retur.filter((r) => r.id !== id) });

  return (
    <div>
      <SectionCard title="Catat Retur" subtitle="Retur 'Sesuai' otomatis menambah stok kembali">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px,1fr))", gap: 10 }}>
          <Field label="Tanggal"><input className="gm-input" type="date" value={form.tanggal} onChange={(e) => setForm({ ...form, tanggal: e.target.value })} /></Field>
          <Field label="Kode Pesanan"><input className="gm-input" value={form.kodePesanan} onChange={(e) => setForm({ ...form, kodePesanan: e.target.value })} /></Field>
          <Field label="Kode Retur"><input className="gm-input" value={form.kodeRetur} onChange={(e) => setForm({ ...form, kodeRetur: e.target.value })} /></Field>
          <Field label="Ekspedisi"><input className="gm-input" value={form.ekspedisi} onChange={(e) => setForm({ ...form, ekspedisi: e.target.value })} /></Field>
          <Field label="Sesuai / Tidak">
            <select className="gm-select" value={form.sesuai} onChange={(e) => setForm({ ...form, sesuai: e.target.value })}>
              <option>Sesuai</option><option>Tidak Sesuai</option>
            </select>
          </Field>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 2fr auto", gap: 10, marginTop: 10, alignItems: "end" }}>
          <Field label="Produk"><ProductSelect products={data.products} value={form.kode} onChange={(v) => setForm({ ...form, kode: v })} /></Field>
          <Field label="Qty"><input className="gm-input" type="number" min="0" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} /></Field>
          <Field label="Catatan"><input className="gm-input" value={form.catatan} onChange={(e) => setForm({ ...form, catatan: e.target.value })} /></Field>
          <button className="gm-btn" onClick={submit}><Plus size={15} /> Simpan</button>
        </div>
      </SectionCard>

      <SectionCard title="Riwayat Retur">
        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th>Tanggal</th><th>Kode Pesanan</th><th>Kode</th><th>Qty</th><th>Status</th><th>Ekspedisi</th><th>Catatan</th><th></th></tr></thead>
            <tbody>
              {data.retur.length === 0 && <EmptyRow colSpan={8} text="Belum ada catatan retur." />}
              {data.retur.map((r) => (
                <tr key={r.id}>
                  <td>{r.tanggal}</td>
                  <td>{r.kodePesanan}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: C.primary }}>{r.kode}</td>
                  <td>{fmt(r.qty)}</td>
                  <td><span style={{ fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 999, background: r.sesuai === "Sesuai" ? "#E4EFE4" : "#F5E1DC", color: r.sesuai === "Sesuai" ? C.success : C.danger }}>{r.sesuai}</span></td>
                  <td>{r.ekspedisi}</td>
                  <td style={{ whiteSpace: "normal", minWidth: 140 }}>{r.catatan}</td>
                  <td><DeleteBtn onClick={() => remove(r.id)} /></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </SectionCard>
    </div>
  );
}

// ---------- Reject ----------
function RejectTab({ data, save, flash }) {
  const emptyForm = { tanggal: todayStr(), kode: "", qty: "", toko: "Shopee", catatan: "" };
  const [form, setForm] = useState(emptyForm);

  const submit = () => {
    if (!form.kode || !Number(form.qty)) { flash("Pilih produk dan isi qty"); return; }
    save({ ...data, reject: [{ id: uid(), ...form, qty: Number(form.qty) }, ...data.reject] });
    setForm({ ...emptyForm, tanggal: form.tanggal });
    flash("Reject dicatat");
  };
  const remove = (id) => save({ ...data, reject: data.reject.filter((r) => r.id !== id) });

  return (
    <div>
      <SectionCard title="Catat Reject" subtitle="Barang rusak / cacat akan mengurangi stok akhir">
        <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr 1fr 1fr auto", gap: 10, alignItems: "end" }}>
          <Field label="Tanggal"><input className="gm-input" type="date" value={form.tanggal} onChange={(e) => setForm({ ...form, tanggal: e.target.value })} /></Field>
          <Field label="Produk"><ProductSelect products={data.products} value={form.kode} onChange={(v) => setForm({ ...form, kode: v })} /></Field>
          <Field label="Qty"><input className="gm-input" type="number" min="0" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} /></Field>
          <Field label="Toko">
            <select className="gm-select" value={form.toko} onChange={(e) => setForm({ ...form, toko: e.target.value })}>
              {TOKO_OPTIONS.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
          </Field>
          <button className="gm-btn" onClick={submit}><Plus size={15} /> Simpan</button>
        </div>
        <div style={{ marginTop: 10 }}>
          <Field label="Catatan"><input className="gm-input" value={form.catatan} onChange={(e) => setForm({ ...form, catatan: e.target.value })} /></Field>
        </div>
      </SectionCard>

      <SectionCard title="Riwayat Reject">
        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th>Tanggal</th><th>Kode</th><th>Produk</th><th>Qty</th><th>Toko</th><th>Catatan</th><th></th></tr></thead>
            <tbody>
              {data.reject.length === 0 && <EmptyRow colSpan={7} text="Belum ada catatan reject." />}
              {data.reject.map((r) => {
                const prod = data.products.find((p) => p.kode === r.kode);
                return (
                  <tr key={r.id}>
                    <td>{r.tanggal}</td>
                    <td style={{ fontFamily: "monospace", fontWeight: 700, color: C.primary }}>{r.kode}</td>
                    <td style={{ whiteSpace: "normal", minWidth: 180 }}>{prod?.nama || "—"}</td>
                    <td>{fmt(r.qty)}</td>
                    <td>{r.toko}</td>
                    <td style={{ whiteSpace: "normal" }}>{r.catatan}</td>
                    <td><DeleteBtn onClick={() => remove(r.id)} /></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </SectionCard>
    </div>
  );
}

// ---------- Pencairan ----------
function PencairanTab({ data, save, flash }) {
  const totalKeluar = data.keluar.reduce((s, k) => s + (Number(k.qty) || 0), 0);
  const totalCair = data.keluar.filter((k) => k.cair).reduce((s, k) => s + (Number(k.qty) || 0), 0);
  const belum = data.keluar.filter((k) => !k.cair);

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px,1fr))", gap: 12, marginBottom: 18 }}>
        <div className="gm-card" style={{ padding: 16 }}>
          <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600, textTransform: "uppercase" }}>Barang Keluar</div>
          <div className="gm-h1" style={{ fontSize: 24, fontWeight: 700, marginTop: 6 }}>{fmt(totalKeluar)}</div>
        </div>
        <div className="gm-card" style={{ padding: 16 }}>
          <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600, textTransform: "uppercase" }}>Sudah Cair</div>
          <div className="gm-h1" style={{ fontSize: 24, fontWeight: 700, marginTop: 6, color: C.success }}>{fmt(totalCair)}</div>
        </div>
        <div className="gm-card" style={{ padding: 16 }}>
          <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600, textTransform: "uppercase" }}>Selisih</div>
          <div className="gm-h1" style={{ fontSize: 24, fontWeight: 700, marginTop: 6, color: C.danger }}>{fmt(totalKeluar - totalCair)}</div>
        </div>
      </div>

      <SectionCard title="Belum Cair" subtitle="Tandai pesanan sebagai 'Sudah Cair' saat dana sudah diterima">
        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th>Tanggal</th><th>Kode Pesanan</th><th>Kode Barang</th><th>Qty</th><th>Toko</th><th></th></tr></thead>
            <tbody>
              {belum.length === 0 && <EmptyRow colSpan={5} text="Semua pesanan sudah dicairkan." />}
              {belum.map((k) => (
                <tr key={k.id}>
                  <td>{k.tanggal}</td>
                  <td>{k.kodePesanan}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: C.primary }}>{k.kode}</td>
                  <td>{fmt(k.qty)}</td>
                  <td>{k.toko}</td>
                  <td>
                    <button className="gm-btn" style={{ padding: "5px 11px", fontSize: 12 }} onClick={() => { save({ ...data, keluar: data.keluar.map((x) => x.id === k.id ? { ...x, cair: true } : x) }); flash("Ditandai sudah cair"); }}>
                      Tandai cair
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </SectionCard>
    </div>
  );
}

// ---------- Master Kode ----------
const KATEGORI = [
  { key: "vendor", label: "Vendor" },
  { key: "jenis", label: "Jenis Barang" },
  { key: "brand", label: "Brand" },
  { key: "bahan", label: "Bahan" },
  { key: "warna", label: "Warna" },
  { key: "ukuran", label: "Ukuran" },
];

function MasterTab({ data, save, flash }) {
  const [cat, setCat] = useState("vendor");
  const [form, setForm] = useState({ code: "", name: "" });
  const list = (data && data.masterCode && data.masterCode[cat]) ? data.masterCode[cat] : [];


  const add = () => {
    if (!form.code.trim() || !form.name.trim()) { flash("Isi kode dan nama"); return; }
    if (list.some((x) => x.code.toUpperCase() === form.code.trim().toUpperCase())) { flash("Kode sudah dipakai di kategori ini"); return; }
    const next = { ...data, masterCode: { ...data.masterCode, [cat]: [...list, { code: form.code.trim().toUpperCase(), name: form.name.trim().toUpperCase() }] } };
    save(next);
    setForm({ code: "", name: "" });
  };
  const remove = (code) => save({ ...data, masterCode: { ...data.masterCode, [cat]: list.filter((x) => x.code !== code) } });

  return (
    <div>
      <SectionCard title="Kode Master" subtitle="Kombinasi kode-kode ini membentuk kode barang otomatis">
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 16 }}>
          {KATEGORI.map((k) => (
            <button
              key={k.key}
              onClick={() => setCat(k.key)}
              style={{
                padding: "7px 13px", borderRadius: 999, fontSize: 12.5, fontWeight: 600, cursor: "pointer",
                border: `1px solid ${cat === k.key ? C.primary : C.border}`,
                background: cat === k.key ? C.primary : "#fff",
                color: cat === k.key ? "#fff" : C.text,
              }}
            >
              {k.label}
            </button>
          ))}
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr auto", gap: 10, marginBottom: 14 }}>
          <Field label="Kode"><input className="gm-input" value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} placeholder="cth: B" /></Field>
          <Field label="Nama"><input className="gm-input" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="cth: TOKO BARU" /></Field>
          <div style={{ display: "flex", alignItems: "end" }}><button className="gm-btn" onClick={add}><Plus size={15} /> Tambah</button></div>
        </div>

        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th>Kode</th><th>Nama</th><th></th></tr></thead>
            <tbody>
              {list.length === 0 && <EmptyRow colSpan={3} text="Belum ada data." />}
              {list.map((v) => (
                <tr key={v.code}>
                  <td style={{ fontFamily: "monospace", fontWeight: 700 }}>{v.code}</td>
                  <td>{v.name}</td>
                  <td><DeleteBtn onClick={() => remove(v.code)} /></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </SectionCard>
    </div>
  );
}

// ---------- Print / PDF report ----------
function PrintTable({ title, rows }) {
  if (!rows.length) {
    return (
      <div style={{ marginBottom: 22 }}>
        <h3 style={{ fontFamily: "Georgia, serif", color: C.primaryDark, fontSize: 15, marginBottom: 4 }}>{title}</h3>
        <div style={{ fontSize: 11, color: C.muted }}>Belum ada data.</div>
      </div>
    );
  }
  const headers = Object.keys(rows[0]);
  return (
    <div style={{ marginBottom: 22, breakInside: "avoid" }}>
      <h3 style={{ fontFamily: "Georgia, serif", color: C.primaryDark, fontSize: 15, marginBottom: 6 }}>{title}</h3>
      <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 10.5 }}>
        <thead>
          <tr>{headers.map((h) => <th key={h} style={{ border: "1px solid #ccc", padding: "4px 7px", background: "#F1E1E1", textAlign: "left" }}>{h}</th>)}</tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i}>{headers.map((h) => <td key={h} style={{ border: "1px solid #ccc", padding: "4px 7px" }}>{String(r[h] ?? "")}</td>)}</tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function PrintReport({ data, stockMap, stokAkhir, onClose }) {
  const { produk, masuk, keluar, retur, reject, masterCode, pencairan } = buildRows(data, stockMap, stokAkhir);
  return (
    <div style={{ background: "#fff", color: "#1a1a1a", fontFamily: "'IBM Plex Sans', Arial, sans-serif", padding: 28, maxWidth: 1000, margin: "0 auto" }}>
      <div className="no-print" style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginBottom: 16 }}>
        <button onClick={onClose} style={{ padding: "8px 14px", fontSize: 13, fontWeight: 600, borderRadius: 8, border: `1px solid ${C.border}`, background: "#fff", cursor: "pointer" }}>Kembali</button>
        <button onClick={() => window.print()} style={{ padding: "8px 14px", fontSize: 13, fontWeight: 600, borderRadius: 8, border: "none", background: C.primary, color: "#fff", cursor: "pointer" }}>Cetak / Simpan PDF</button>
      </div>
      <h1 style={{ fontFamily: "Georgia, serif", color: C.primaryDark, marginBottom: 0 }}>Pembukuan Gamis Niki Dziyab</h1>
      <div style={{ fontSize: 11.5, color: C.muted, marginBottom: 20 }}>Dicetak: {todayStr()}</div>
      <PrintTable title="Produk & Stok" rows={produk} />
      <PrintTable title="Barang Masuk" rows={masuk} />
      <PrintTable title="Barang Keluar" rows={keluar} />
      <PrintTable title="Retur" rows={retur} />
      <PrintTable title="Reject" rows={reject} />
      <PrintTable title="Pencairan" rows={pencairan} />
      <PrintTable title="Kode Master" rows={masterCode} />
      <style>{`
        @media print {
          .no-print { display: none !important; }
          @page { margin: 14mm; }
        }
      `}</style>
    </div>
  );
}
