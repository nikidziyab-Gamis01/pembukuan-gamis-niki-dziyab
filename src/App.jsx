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
  Users,
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

const STORAGE_KEY = "gamis-inventory-v1";
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
    gajiDinamis: source.gajiDinamis && typeof source.gajiDinamis === "object" ? source.gajiDinamis : {},
    gajiSetelan: {
      gajiPokokOwner: Number(source.gajiSetelan?.gajiPokokOwner) || 0,
      komisiOwnerPerPcs: Number(source.gajiSetelan?.komisiOwnerPerPcs) || 0,
      komisiHostPerPcs: Number(source.gajiSetelan?.komisiHostPerPcs) || 0,
      komisiAdminPerPcs: Number(source.gajiSetelan?.komisiAdminPerPcs) || 0,
    },
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
    Kode: p.kode, "Nama Barang": p.nama, Warna: p.warna, Ukuran: p.ukuran, Vendor: p.vendor, "HPP Modal Master": p.hpp || 0, "Stok Akhir": stokAkhir(p.kode)
  }));

  const masuk = (data.masuk || []).map((m) => {
    const matchedProd = (data.products || []).find((p) => p.kode === m.kode) || {};
    return {
      Tanggal: m.tanggal, "Kode Barang": m.kode, "Nama Barang": matchedProd.nama || "—", Qty: m.qty, Vendor: m.vendor, "HPP Satuan": m.hpp || matchedProd.hpp || 0, "Total Modal Belanja": (m.qty * (m.hpp || matchedProd.hpp || 0)), Catatan: m.catatan
    };
  });

  const keluar = (data.keluar || []).map((k) => {
    const matchedProd = (data.products || []).find((p) => p.kode === k.kode) || {};
    return {
      Tanggal: k.tanggal, "Kode Pesanan": k.kodePesanan, Shift: k.shift, Host: k.host,
      "Kode Barang": k.kode, "Nama Barang": matchedProd.nama || "—", Qty: k.qty, Toko: k.toko, "Harga Jual Aplikasi": k.hargaJual || 0, "Omset Bruto": (k.qty * (k.hargaJual || 0)), "Nominal Pencairan Bersih": k.nominalCairRiil || 0, "COD/NON COD": k.cod, Catatan: k.catatan || "", "Sudah Cair": k.cair ? "Ya" : "Belum",
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
  const totalCair = (data.keluar || []).filter((k) => k.cair).reduce((s, k) => s + (Number(k.nominalCairRiil) || 0), 0);
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
  const headers = Object.keys(rows);
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
  { id: "labaRugi", label: "📈 Laporan Laba Rugi", icon: PieChart },
  { id: "bukuKas", label: "📊 Master Buku Kas", icon: PieChart },
  { id: "keuangan", label: "💰 Keuangan & Gaji Tim", icon: Settings2 },
  { id: "master", label: "⚙️ Kode Master", icon: Settings2 },
];

export default function GamisInventoryApp() {
  const [data, setData] = useState(EMPTY_DATA);
  const [tab, setTab] = useState("dashboard");
  const [toast, setToast] = useState("");
  const [printView, setPrintView] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [user, setUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [authError, setAuthError] = useState("");
  const fileInputRef = useRef(null);

  const loadWhitelistEmails = async () => {
    const whitelistRef = doc(db, "pembukuan_gamis", "whitelist_users");
    const whitelistSnap = await getDoc(whitelistRef);
    let allowedEmails = ["nikidziyab@gmail.com"];
    if (whitelistSnap.exists()) {
      const allowedData = whitelistSnap.data();
      if (allowedData && Array.isArray(allowedData.emails)) {
        allowedEmails = [...allowedEmails, ...allowedData.emails];
      }
    }
    return allowedEmails.map((e) => e.toLowerCase());
  };

  useEffect(() => {
    const unsubscribeAuth = onAuthStateChanged(auth, async (currentUser) => {
      if (currentUser) {
        try {
          if (currentUser.email.toLowerCase() === "nikidziyab@gmail.com") {
            setUser(currentUser);
            setAuthLoading(false);
            return;
          }
          const whitelistRef = doc(db, "pembukuan_gamis", "whitelist_users");
          const whitelistSnap = await getDoc(whitelistRef);
          let allowedEmails = [];
          if (whitelistSnap.exists()) {
            const allowedData = whitelistSnap.data();
            if (allowedData && Array.isArray(allowedData.emails)) {
              allowedEmails = allowedData.emails.map((e) => String(e || "").toLowerCase());
            }
          }
          if (allowedEmails.includes(currentUser.email.toLowerCase())) {
            setUser(currentUser);
          } else {
            alert(`Akses Ditolak! Alamat Email (${currentUser.email}) bukan bagian dari Tim Resmi Niki Dziyab.`);
            signOut(auth);
            setUser(null);
          }
        } catch (error) {
          console.error("Gagal memvalidasi daftar email:", error);
          setUser(null);
        }
      } else {
        setUser(null);
      }
      setAuthLoading(false);
    });
    return () => unsubscribeAuth();
  }, []);

  const loginWithGoogle = async () => {
    setAuthError("");
    try {
      const result = await signInWithPopup(auth, googleProvider);
      const safeEmail = result.user?.email || result.user?.providerData?.[0]?.email || "";
      const allowedEmails = await loadWhitelistEmails();
      if (!allowedEmails.includes(safeEmail.toLowerCase())) {
        alert(`Akses Ditolak! Alamat Email (${safeEmail || "tidak tersedia"}) bukan bagian dari Tim Resmi Niki Dziyab.`);
        await signOut(auth);
        setUser(null);
        return;
      }
      setUser(result.user);
    } catch (error) {
      console.error("Gagal melakukan autentikasi Google:", error);
      const code = error?.code || "";
      if (code === "auth/popup-closed-by-user") {
        setAuthError("Login dibatalkan. Silakan klik tombol Google kembali.");
      } else if (code === "auth/popup-blocked") {
        setAuthError("Popup diblokir browser. Izinkan popup untuk aplikasi ini lalu coba lagi.");
      } else {
        setAuthError("Login Google gagal. Cek konfigurasi Firebase dan izin browser.");
      }
    }
  };

  const logout = async () => {
    try { await signOut(auth); } catch (e) { console.error(e); }
  };

  // ----- KODE PENGGANTI MULTI-ACCOUNT DATABASE (FINAL) -----
  useEffect(() => {
    if (!user) return;

    // Otomatis membagi target dokumen berdasarkan email login
    const userEmail = user.email.toLowerCase();
    const documentTarget = userEmail === "nikidziyab@gmail.com" ? "data_utama" : "data_utama_teman";

    const unsubscribe = onSnapshot(doc(db, "pembukuan_gamis", documentTarget), (snapshot) => {
      if (snapshot.exists()) {
        const cloudData = snapshot.data();
        const payload = cloudData && cloudData.pembukuanData ? cloudData.pembukuanData : EMPTY_DATA;
        setData(normalizeData(payload));
      } else {
        setData(EMPTY_DATA);
      }
    });

    return () => unsubscribe();
  }, [user]);

  const syncToCloud = async (nextData = data) => {
    if (!user) return false;
    const normalized = normalizeData(nextData);
    setData(normalized);

    const userEmail = user.email.toLowerCase();
    const documentTarget = userEmail === "nikidziyab@gmail.com" ? "data_utama" : "data_utama_teman";

    try {
      const docRef = doc(db, "pembukuan_gamis", documentTarget);
      await setDoc(docRef, { pembukuanData: normalized, updatedAt: new Date() });
      return true;
    } catch (e) {
      console.error("Gagal menyimpan data ke Firebase:", e);
      return false;
    }
  };

  const save = async (next) => {
    const normalized = normalizeData(next);
    return syncToCloud(normalized);
  };
  // ----- END KODE PENGGANTI -----

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

  const saveAsFile = () => {
    try {
      const payload = JSON.stringify(normalizeData(data), null, 2);
      downloadBlob(payload, `pembukuan-gamis-backup-${todayStr()}.json`, "application/json");
      flash("File backup berhasil diunduh");
    } catch (e) {
      flash("Gagal menyimpan file cadangan");
    }
  };

  const syncCloudNow = async () => {
    const ok = await syncToCloud(data);
    flash(ok ? "Data berhasil disinkronkan ke cloud" : "Sinkronisasi gagal");
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
    data.products.forEach((p) => { map[p.kode] = { masuk: 0, keluar: 0, returOk: 0, reject: 0 }; });
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

  if (authLoading) {
    return (
      <div style={{ display: "flex", justifyContent: "center", alignItems: "center", minHeight: "100vh", backgroundColor: "#f8f9fa", fontFamily: "sans-serif" }}>
        <div style={{ color: "#7c2d12", fontSize: 16, fontWeight: "bold" }}>Mengamankan jaringan pembukuan...</div>
      </div>
    );
  }

  if (!user) {
    return (
      <div style={{ display: "flex", flexDirection: "column", justifyContent: "center", alignItems: "center", minHeight: "100vh", backgroundColor: "#f8f9fa", fontFamily: "sans-serif" }}>
        <div style={{ backgroundColor: "#fff", padding: 40, borderRadius: 12, boxShadow: "0 4px 12px rgba(0,0,0,0.08)", textAlign: "center", maxWidth: 400, width: "90%" }}>
          <h2 style={{ color: "#7c2d12", marginBottom: 10, fontSize: 24, fontWeight: "bold" }}>Pembukuan Niki Dziyab</h2>
          <p style={{ color: "#6b7280", fontSize: 14, lineHeight: "1.5", marginBottom: 30 }}>Sistem Keamanan Aktif. Akses dibatasi khusus internal. Silakan masuk menggunakan email tim operasional Anda yang sah.</p>
          {authError && (
            <div style={{ background: "#fff1f2", border: "1px solid #fecdd3", color: "#9f1239", borderRadius: 8, padding: "10px 12px", marginBottom: 16, fontSize: 12.5, textAlign: "left" }}>
              {authError}
            </div>
          )}
          <button onClick={loginWithGoogle} style={{ backgroundColor: "#10b981", color: "#fff", border: "none", padding: "12px 24px", borderRadius: 8, fontSize: 15, fontWeight: "bold", cursor: "pointer", display: "flex", alignItems: "center", gap: 10, margin: "0 auto" }}>
            🔑 Masuk dengan Google
          </button>
        </div>
      </div>
    );
  }

  if (printView) {
    return <PrintReport data={data} stockMap={stockMap} stokAkhir={stokAkhir} onClose={() => setPrintView(false)} />;
  }
  return (
    <div style={{ background: C.bg, minHeight: 640, fontFamily: "'IBM Plex Sans', sans-serif", color: C.text }}>
      <style>{`
        @import url('https://googleapis.com');
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
                <div style={{ fontSize: 11.5, color: C.muted }}>Stok & HPP Master · Barang Masuk & Keluar · Retur · Reject · Pencairan · Keuangan</div>
              </div>
            </div>

            <div style={{ display: "flex", alignItems: "center", gap: 6, position: "relative", flexWrap: "wrap", justifyContent: "flex-end" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, background: C.surfaceAlt, border: `1px solid ${C.border}`, borderRadius: 999, padding: "6px 10px 6px 8px", marginRight: 4 }}>
                {user?.photoURL ? (
                  <img src={user.photoURL} alt="Avatar" style={{ width: 26, height: 26, borderRadius: "50%", objectFit: "cover", border: `1px solid ${C.border}` }} />
                ) : (
                  <div style={{ width: 26, height: 26, borderRadius: "50%", background: C.primary, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 700 }}>
                    {((user?.displayName || user?.email || "A").charAt(0) || "A").toUpperCase()}
                  </div>
                )}
                <div style={{ display: "flex", flexDirection: "column", lineHeight: 1.2, minWidth: 0 }}>
                  <span style={{ fontSize: 11, color: C.muted, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".04em" }}>Akun</span>
                  <span style={{ fontSize: 11.5, color: C.primaryDark, fontWeight: 600, maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                    {user?.displayName || user?.email || "Google User"}
                  </span>
                </div>
              </div>

              <button className="gm-btn-ghost" style={{ borderRadius: 8, padding: "8px 12px", fontSize: 12.5, fontWeight: 600, display: "inline-flex", alignItems: "center", gap: 6, cursor: "pointer", background: "transparent" }} onClick={logout}>
                <XCircle size={14} /> Logout
              </button>

              <input ref={fileInputRef} type="file" accept="application/json" onChange={handleImportFile} style={{ display: "none" }} />
              <button className="gm-btn-ghost" style={{ borderRadius: 8, padding: "8px 12px", fontSize: 12.5, fontWeight: 600, display: "inline-flex", alignItems: "center", gap: 6, cursor: "pointer", background: "transparent" }} onClick={() => fileInputRef.current?.click()}>
                <Upload size={14} /> Impor
              </button>
              <button className="gm-btn-ghost" style={{ borderRadius: 8, padding: "8px 12px", fontSize: 12.5, fontWeight: 600, display: "inline-flex", alignItems: "center", gap: 6, cursor: "pointer", background: "transparent" }} onClick={saveAsFile}>
                <Download size={14} /> Backup File
              </button>
              <button className="gm-btn" style={{ padding: "8px 14px", fontSize: 12.5, background: C.success }} onClick={syncCloudNow}>
                <Save size={14} /> Update Real-time
              </button>
              <div style={{ position: "relative" }}>
                <button className="gm-btn" style={{ padding: "8px 14px", fontSize: 12.5, background: C.accent }} onClick={() => setMenuOpen((v) => !v)}>
                  Export {menuOpen ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
                </button>
                {menuOpen && (
                  <div style={{ position: "absolute", right: 0, top: "calc(100% + 6px)", background: "#fff", border: `1px solid ${C.border}`, borderRadius: 10, boxShadow: "0 10px 28px rgba(0,0,0,.14)", overflow: "hidden", zIndex: 30, minWidth: 190 }}>
                    <MenuItem icon={FileSpreadsheet} label="Export ke Excel (.xlsx)" onClick={() => { exportExcel(data, stockMap, stokAkhir); setMenuOpen(false); flash("Excel diunduh"); }} />
                    <MenuItem icon={FileText} label="Export ke Word (.doc)" onClick={() => { exportWord(data, stockMap, stokAkhir); setMenuOpen(false); flash("Word diunduh"); }} />
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
        {tab === "labaRugi" && <LabaRugiTab data={data} />}
        {tab === "bukuKas" && <BukuKasTab data={data} save={save} flash={flash} />}
        {tab === "keuangan" && <KeuanganTab data={data} save={save} flash={flash} />}
        {tab === "master" && <MasterTab data={data} save={save} flash={flash} />}
      </div>

           {/* FOOTER HAK CIPTA GLOBAL (MUNCUL DI SEMUA HALAMAN DAN TAB KERJA APLIKASI) */}
      <div style={{ textAlign: "center", padding: "24px 0 20px", borderTop: `1px solid ${C.border}`, marginTop: 40, background: C.surface, width: "100%" }}>
        <p style={{ fontSize: "12px", color: C.muted, letterSpacing: "0.02em", margin: 0, fontWeight: 500 }}>
          © 2026 <span style={{ color: C.primary, fontWeight: 700 }}>Niki Dziyab Omnichannel System</span>. All Rights Reserved.
        </p>
        <p style={{ fontSize: "11px", color: C.muted, marginTop: 4, margin: 0 }}>
          Product developed with pride by <a href="https://syahrulabs.id" target="_blank" rel="noreferrer" style={{ color: C.secondary, fontWeight: 600, textDecoration: "none" }}>Syahrulabs.id</a>
        </p>
      </div>

      {toast && (
        <div style={{ position: "fixed", bottom: 20, left: "50%", transform: "translateX(-50%)", background: C.primaryDark, color: "#fff", padding: "10px 18px", borderRadius: 999, fontSize: 13, fontWeight: 500, boxShadow: "0 8px 24px rgba(0,0,0,.2)" }}>
          {toast}
        </div>
      )}
    </div>
  );
}


// ---------- REFIX FINAL: MEMBERSIHKAN KOMPONEN FIELD GLOBAL ----------
function Field({ label, children }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
      <label style={{ fontSize: "11.5px", fontWeight: 600, color: C.muted, textTransform: "uppercase", letterSpacing: ".03em" }}>
        {label}
      </label>
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

// ---------- DASHBOARD UTUH DENGAN OPTIMASI GRAFIK TOP SELLING HORIZONTAL ----------
function Dashboard({ data, stockMap, stokAkhir }) {
  const safeData = normalizeData(data || {});
  const totalProduk = safeData.products?.length || 0;
  const totalStok = (safeData.products || []).reduce((s, p) => s + stokAkhir(p.kode), 0);
  const bulanIni = monthOf(todayStr());
  const keluarBulanIni = (safeData.keluar || []).filter((k) => monthOf(k.tanggal) === bulanIni).reduce((s, k) => s + (Number(k.qty) || 0), 0);
  const returBulanIni = (safeData.retur || []).filter((r) => monthOf(r.tanggal) === bulanIni).reduce((s, r) => s + (Number(r.qty) || 0), 0);
  const rejectBulanIni = (safeData.reject || []).filter((r) => monthOf(r.tanggal) === bulanIni).reduce((s, r) => s + (Number(r.qty) || 0), 0);

  const totalDebit = (safeData.bukuKas || []).filter((b) => b.jenis === "Debit").reduce((s, b) => s + (Number(b.nominal) || 0), 0);
  const totalKredit = (safeData.bukuKas || []).filter((b) => b.jenis === "Kredit").reduce((s, b) => s + (Number(b.nominal) || 0), 0);
  const saldoKasRiil = totalDebit - totalKredit;

  const totalKeluar = (safeData.keluar || []).reduce((s, k) => s + (Number(k.qty) || 0), 0);
  const totalCair = (safeData.keluar || []).filter((k) => k.cair).reduce((s, k) => s + (Number(k.nominalCairRiil) || 0), 0);

  // Perhitungan Laba Kotor & HPP Terjual (Hanya jualan yang berstatus CAIR RIIL)
  const barangCairList = (safeData.keluar || []).filter((k) => k.cair);
  const totalHppTerjual = barangCairList.reduce((sum, k) => {
    const prod = (safeData.products || []).find((p) => p.kode === k.kode) || {};
    return sum + ((Number(k.qty) || 0) * (prod.hpp || 0));
  }, 0);
  const labaRugiBersihPencairan = totalCair - totalHppTerjual;

  // Analisis 5 Produk Terlaris (Top Selling Products)
  const productSalesMap = {};
  (safeData.keluar || []).forEach((k) => {
    productSalesMap[k.kode] = (productSalesMap[k.kode] || 0) + (Number(k.qty) || 0);
  });
  const topProductsData = Object.entries(productSalesMap)
    .map(([kode, qty]) => {
      const prod = (safeData.products || []).find((p) => p.kode === kode);
      return { sku: kode, nama: prod ? prod.nama : kode, qty };
    })
    .sort((a, b) => b.qty - a.qty)
    .slice(0, 5);

  const chartData = BULAN_ID.slice(6, 12).map((b) => ({
    bulan: b.slice(0, 3),
    qty: (safeData.keluar || []).filter((k) => monthOf(k.tanggal) === b).reduce((s, k) => s + (Number(k.qty) || 0), 0),
  }));

  const lowStock = (safeData.products || []).filter((p) => stokAkhir(p.kode) <= 10);

  const cards = [
    { label: "Total Ragam Produk", value: fmt(totalProduk), color: C.primary },
    { label: "Stok Tersedia Gudang", value: fmt(totalStok), color: C.success },
    { label: `Keluar — ${bulanIni}`, value: fmt(keluarBulanIni), color: C.accent },
    { label: "Saldo Buku Kas Toko", value: `Rp ${fmt(saldoKasRiil)}`, color: saldoKasRiil >= 0 ? C.success : C.danger },
    { label: "Total HPP Terjual (Cair)", value: `Rp ${fmt(totalHppTerjual)}`, color: C.secondary },
    { label: "Estimasi Laba Bersih", value: `Rp ${fmt(labaRugiBersihPencairan)}`, color: labaRugiBersihPencairan >= 0 ? C.success : C.danger },
  ];

  return (
    <div>
      {/* RENDER KOTAK METRIK UTAMA */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 12, marginBottom: 18 }}>
        {cards.map((c) => (
          <div key={c.label} className="gm-card" style={{ padding: 16 }}>
            <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600, textTransform: "uppercase", letterSpacing: ".03em" }}>{c.label}</div>
            <div className="gm-h1" style={{ fontSize: 22, fontWeight: 700, color: c.color, marginTop: 6 }}>{c.value}</div>
          </div>
        ))}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1.2fr 1.3fr", gap: 16, alignItems: "stretch", marginBottom: 18 }}>
        {/* GRAFIK PENJUALAN BULANAN */}
        <SectionCard title="Tren Volume Barang Keluar" subtitle="Analisis Kuantitas Jualan Bulanan">
          <div style={{ height: 220 }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartData} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={C.border} vertical={false} />
                <XAxis dataKey="bulan" tick={{ fontSize: 12, fill: C.muted }} axisLine={{ stroke: C.border }} tickLine={false} />
                <YAxis tick={{ fontSize: 12, fill: C.muted }} axisLine={false} tickLine={false} />
                <Tooltip contentStyle={{ borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 12 }} />
                <Bar dataKey="qty" fill={C.primary} radius={[5, 5, 0, 0]}>
                  {chartData.map((entry, i) => (
                    <Cell key={i} fill={entry.bulan === bulanIni.slice(0, 3) ? C.primary : C.accentSoft} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </SectionCard>

        {/* GRAFIK FIX FINAL: TOP SELLING HORIZONTAL LAYOUT (TEXT SKU AMAN TEGAK DI ATAS BALOK) */}
        <SectionCard title="🔥 5 Produk Terlaris (Top Selling)" subtitle="Berdasarkan Kuantitas Unit Terjual">
          <div style={{ height: 220 }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={topProductsData} layout="vertical" margin={{ top: 15, right: 25, left: -30, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={C.border} horizontal={false} />
                <XAxis type="number" tick={{ fontSize: 11, fill: C.muted }} axisLine={false} />
                <YAxis dataKey="sku" type="category" hide={true} />
                <Tooltip contentStyle={{ fontSize: 11, borderRadius: 6 }} />
                <Bar 
                  dataKey="qty" 
                  fill={C.accent} 
                  radius={[0, 4, 4, 0]} 
                  barSize={14} 
                  label={{ position: "top", dataKey: "sku", fill: C.primaryDark, fontSize: 10, fontWeight: 700 }} 
                />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </SectionCard>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 18 }}>
        <SectionCard title="Ringkasan Arus Dana Masuk" subtitle="Volume keluar vs konfirmasi dana cair bersih">
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <Stat label="Total Volume Keluar Toko" value={`${fmt(totalKeluar)} Pcs`} />
            <Stat label="Total Dana Sukses Cair Riil" value={`Rp ${fmt(totalCair)}`} color={C.success} />
            <Stat label="Selisih Dana Belum Cair" value={`Rp ${fmt((safeData.keluar || []).filter(k => !k.cair).reduce((s, k) => s + (k.qty * (k.hargaJual || 0)), 0))}`} color={C.danger} />
          </div>
        </SectionCard>

        <SectionCard title="Stok Menipis" subtitle="Produk dengan sisa unit ≤ 10" right={lowStock.length > 0 && <AlertTriangle size={16} color={C.danger} />}>
          {lowStock.length === 0 ? (
            <div style={{ fontSize: 13, color: C.muted }}>Seluruh stok aman berada di atas batas minimum.</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 115, overflowY: "auto", paddingRight: 4 }}>
              {lowStock.map((p) => (
                <div key={p.kode} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "6px 10px", background: C.surfaceAlt, borderRadius: 8 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ width: 8, height: 8, borderRadius: "50%", background: WARNA_SWATCH[p.warna] || C.muted, display: "inline-block" }} />
                    <span style={{ fontSize: 12, fontWeight: 500 }}>{p.kode} — {p.nama}</span>
                  </div>
                  <span style={{ fontSize: 12, fontWeight: 700, color: C.danger }}>{stokAkhir(p.kode)} Pcs</span>
                </div>
              ))}
            </div>
          )}
        </SectionCard>
      </div>
    </div>
  );
}

// ---------- FUNGSI STAT PENDUKUNG DASBOR ----------
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
  const safeData = normalizeData(data || {});
  const products = safeData.products || [];
  const mc = safeData.masterCode || EMPTY_MASTER;
  const [form, setForm] = useState({ vendor: "", jenis: "", brand: "", bahan: "", warna: "", ukuran: "", nama: "", hpp: "", stokAwal: "" });

  const codeOf = (list, name) => (Array.isArray(list) ? (list.find((x) => x.name === name) || {}).code : "") || "";
  const previewKode = useMemo(() => {
    if (form.vendor && form.jenis && form.brand && form.bahan && form.warna && form.ukuran) {
      return codeOf(mc.vendor, form.vendor) + codeOf(mc.jenis, form.jenis) + codeOf(mc.brand, form.brand) + codeOf(mc.bahan, form.bahan) + codeOf(mc.warna, form.warna) + codeOf(mc.ukuran, form.ukuran);
    }
    return "";
  }, [form, mc]);

  const addProduk = () => {
    if (!previewKode) { flash("Lengkapi vendor, jenis, brand, bahan, warna, dan ukuran dulu"); return; }
    if (products.some((p) => p.kode === previewKode)) { flash("Kode barang sudah terdaftar"); return; }
    if (!form.hpp || Number(form.hpp) <= 0) { flash("Wajib mengisi Harga HPP sebagai nilai modal master!"); return; }

    const nama = form.nama.trim() || `${form.brand} ${form.jenis} ${form.warna} ${form.ukuran}`.toUpperCase();
    const next = { ...safeData, products: [...products, { kode: previewKode, nama, jenis: form.jenis, brand: form.brand, bahan: form.bahan, warna: form.warna, ukuran: form.ukuran, vendor: form.vendor, hpp: Number(form.hpp) }] };
    
    if (Number(form.stokAwal) > 0) {
      next.masuk = [{ id: uid(), tanggal: todayStr(), kode: previewKode, qty: Number(form.stokAwal), vendor: form.vendor, hpp: Number(form.hpp), catatan: "Stok awal" }, ...(safeData.masuk || [])];
    }
    save(next);
    setForm({ vendor: "", jenis: "", brand: "", bahan: "", warna: "", ukuran: "", nama: "", hpp: "", stokAwal: "" });
    flash(`Produk ${previewKode} ditambahkan`);
  };

  const removeProduk = (kode) => {
    const hasLogs = [...(safeData.masuk || []), ...(safeData.keluar || []), ...(safeData.retur || []), ...(safeData.reject || [])].some((r) => r.kode === kode);
    if (hasLogs) { flash("Tidak bisa dihapus — produk ini punya riwayat transaksi mutasi"); return; }
    save({ ...safeData, products: products.filter((p) => p.kode !== kode) });
    flash("Produk berhasil dihapus");
  };

  const filtered = products.filter((p) => {
    if (!p) return false;
    const kataKunci = q.trim().toLowerCase();
    if (!kataKunci) return true;
    return String(p.kode || "").toLowerCase().includes(kataKunci) || String(p.nama || "").toLowerCase().includes(kataKunci) || String(p.warna || "").toLowerCase().includes(kataKunci);
  });
  return (
    <div>
      <SectionCard title="Tambah Ragam Produk & Harga HPP" subtitle="Kode SKU dikonstruksi otomatis berdasarkan kriteria master code terpilih">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px,1fr))", gap: 10 }}>
          <Field label="Vendor Konveksi">
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
          <Field label="Brand Pakaian">
            <select className="gm-select" value={form.brand} onChange={(e) => setForm({ ...form, brand: e.target.value })}>
              <option value="">Pilih…</option>
              {(mc.brand || []).map((v) => <option key={v.code} value={v.name}>{v.name}</option>)}
            </select>
          </Field>
          <Field label="Bahan Kain">
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
          <Field label="Ukuran / Size">
            <select className="gm-select" value={form.ukuran} onChange={(e) => setForm({ ...form, ukuran: e.target.value })}>
              <option value="">Pilih…</option>
              {(mc.ukuran || []).map((v) => <option key={v.code} value={v.name}>{v.name}</option>)}
            </select>
          </Field>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1.5fr 1fr 1fr auto", gap: 10, marginTop: 10, alignItems: "end" }}>
          <Field label="Nama Deskripsi Unit (Otomatis jika kosong)">
            <input className="gm-input" value={form.nama} onChange={(e) => setForm({ ...form, nama: e.target.value })} placeholder="Nama barang lengkap" />
          </Field>
          <Field label="Harga HPP Master Modal (Rp)">
            <input className="gm-input" type="number" value={form.hpp} onChange={(e) => setForm({ ...form, hpp: e.target.value })} placeholder="0" />
          </Field>
          <Field label="Kuantitas Awal">
            <input className="gm-input" type="number" min="0" value={form.stokAwal} onChange={(e) => setForm({ ...form, stokAwal: e.target.value })} placeholder="0" />
          </Field>
          <button className="gm-btn" onClick={addProduk}><Plus size={15} /> Daftarkan SKU</button>
        </div>
        {previewKode && (
          <div style={{ marginTop: 10, fontSize: 12.5, color: C.muted }}>
            Kode SKU Terbentuk: <span style={{ fontWeight: 700, color: C.primary, fontFamily: "monospace", fontSize: 13.5 }}>{previewKode}</span>
          </div>
        )}
      </SectionCard>

      <SectionCard title="Daftar Manajemen Inventaris Stok & Nilai HPP" right={
        <div style={{ position: "relative", width: 220 }}>
          <Search size={14} style={{ position: "absolute", left: 9, top: 10, color: C.muted }} />
          <input className="gm-input" style={{ paddingLeft: 28 }} placeholder="Cari kode / nama / warna" value={q} onChange={(e) => setQ(e.target.value)} />
        </div>
      }>
        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th>Kode SKU</th><th>Nama Unit</th><th>Varian Warna</th><th>Size</th><th>Harga HPP</th><th>Masuk</th><th>Terjual</th><th>Stok Akhir</th><th>Aksi</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 && <EmptyRow colSpan={9} text="Tidak ditemukan data produk terkait." />}
              {filtered.map((p) => {
                const s = stockMap[p.kode] || { masuk: 0, keluar: 0, returOk: 0, reject: 0 };
                const akhir = stokAkhir(p.kode);
                return (
                  <tr key={p.kode}>
                    <td style={{ fontFamily: "monospace", fontWeight: 700, color: C.primary }}>{p.kode}</td>
                    <td style={{ whiteSpace: "normal" }}>{p.nama}</td>
                    <td>{p.warna}</td>
                    <td>{p.ukuran}</td>
                    <td style={{ fontWeight: 600 }}>Rp {fmt(p.hpp || 0)}</td>
                    <td>{fmt(s.masuk)}</td>
                    <td>{fmt(s.keluar)}</td>
                    <td style={{ fontWeight: 700, color: akhir <= 10 ? C.danger : C.success }}>{fmt(akhir)} Pcs</td>
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
  const safeData = normalizeData(data || {});
  const [form, setForm] = useState({ tanggal: todayStr(), kode: "", qty: "", hpp: "", catatan: "" });

  const produkTerpilih = (safeData.products || []).find((p) => p.kode === form.kode);

  useEffect(() => {
    if (produkTerpilih) {
      setForm((f) => ({ ...f, hpp: produkTerpilih.hpp || "" }));
    }
  }, [form.kode, produkTerpilih]);

  const submit = () => {
    if (!form.kode || !Number(form.qty)) { flash("Pilih produk dan isi kuantitas"); return; }
    
    const hppFinal = Number(form.hpp) || produkTerpilih?.hpp || 0;
    const entry = { id: uid(), tanggal: form.tanggal, kode: form.kode, qty: Number(form.qty), hpp: hppFinal, vendor: produkTerpilih?.vendor || "", catatan: form.catatan };
    
    save({ ...safeData, masuk: [entry, ...(safeData.masuk || [])] });
    setForm({ tanggal: form.tanggal, kode: "", qty: "", hpp: "", catatan: "" });
    flash("Log restock barang masuk berhasil dicatat");
  };

  const remove = (id) => {
    if (confirm("Hapus catatan barang masuk ini?")) {
      save({ ...safeData, masuk: (safeData.masuk || []).filter((m) => m.id !== id) });
      flash("Catatan berhasil dihapus");
    }
  };

  return (
    <div>
      <SectionCard title="Pencatatan Log Masuk Persediaan" subtitle="Kolom nominal HPP otomatis memuat harga modal master, namun tetap dapat diubah manual jika terjadi fluktuasi biaya produksi dari vendor">
        <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr 1fr 1fr 1.5fr auto", gap: 10, alignItems: "end" }}>
          <Field label="Tanggal"><input className="gm-input" type="date" value={form.tanggal} onChange={(e) => setForm({ ...form, tanggal: e.target.value })} /></Field>
          <Field label="Produk SKU"><ProductSelect products={safeData.products || []} value={form.kode} onChange={(v) => setForm({ ...form, kode: v })} /></Field>
          <Field label="Qty Restock"><input className="gm-input" type="number" min="0" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} /></Field>
          <Field label="Nominal HPP Masuk (Rp)"><input className="gm-input" type="number" value={form.hpp} onChange={(e) => setForm({ ...form, hpp: e.target.value })} placeholder="Bawaan Master" /></Field>
          <Field label="Catatan / Batch"><input className="gm-input" value={form.catatan} onChange={(e) => setForm({ ...form, catatan: e.target.value })} placeholder="Opsional" /></Field>
          <button className="gm-btn" onClick={submit}><Plus size={15} /> Simpan Nota</button>
        </div>
      </SectionCard>

      <SectionCard title="Riwayat Mutasi Barang Masuk">
        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th>Tanggal</th><th>Kode SKU</th><th>Nama Barang</th><th>Qty</th><th>HPP Satuan</th><th>Total Nilai Pembelian</th><th>Vendor</th><th>Catatan</th><th></th></tr></thead>
            <tbody>
              {(safeData.masuk || []).length === 0 && <EmptyRow colSpan={9} text="Belum ada catatan pasokan barang masuk." />}
              {(safeData.masuk || []).map((m) => {
                const prod = (safeData.products || []).find((p) => p.kode === m.kode);
                const hppTerapkan = m.hpp || prod?.hpp || 0;
                return (
                  <tr key={m.id}>
                    <td>{m.tanggal}</td>
                    <td style={{ fontFamily: "monospace", fontWeight: 700, color: C.primary }}>{m.kode}</td>
                    <td style={{ whiteSpace: "normal" }}>{prod?.nama || "—"}</td>
                    <td>{fmt(m.qty || 0)} Pcs</td>
                    <td>Rp {fmt(hppTerapkan)}</td>
                    <td style={{ fontWeight: 700, color: C.secondary }}>Rp {fmt((m.qty || 0) * hppTerapkan)}</td>
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

// ---------- Barang Keluar (Omnichannel, Pencairan Manual & Skema Ongkir Sample Affiliate) ----------
function KeluarTab({ data, save, stokAkhir, flash }) {
  const safeData = normalizeData(data || {});
  const emptyForm = { tanggal: todayStr(), kodePesanan: "", shift: "", host: "", kode: "", qty: "", toko: "Shopee", hargaJual: "", cod: "NON COD", affiliate: false, namaAffiliate: "", pengiriman: "", ongkir: "", catatan: "" };
  const [form, setForm] = useState(emptyForm);
  const [monthFilter, setMonthFilter] = useState("SEMUA");
  const [inputCairManual, setInputCairManual] = useState({});
  
  // State bantuan untuk mencatat biaya ongkir manual khusus Sample Affiliate
  const [ongkirSampleManual, setOngkirSampleManual] = useState({});

  const karyawanList = (safeData.karyawanMaster || []).map((k) => k.nama).filter(Boolean) || [];
  const isSampleCategory = form.toko === "Sample Affiliate (Aplikasi)" || form.toko === "Sample Affiliate (Manual + Ongkir Toko)";

  const submit = () => {
    if (!form.kode || !Number(form.qty)) { flash("Pilih produk dan isi qty keluar!"); return; }
    
    // Validasi harga jual jika bukan kategori Sample Affiliate
    if (!isSampleCategory && !form.hargaJual) {
      flash("Mohon isi harga jual bruto aplikasi!");
      return;
    }

    const sisa = stokAkhir(form.kode);
    if (Number(form.qty) > sisa) { flash(`Stok tidak cukup (tersisa ${sisa} Pcs)`); return; }
    
    // Jika Sample Affiliate, harga jual bruto otomatis diset Rp 0
    const hargaJualFinal = isSampleCategory ? 0 : Number(form.hargaJual);

    const entry = { 
      id: uid(), 
      ...form, 
      qty: Number(form.qty), 
      hargaJual: hargaJualFinal, 
      cair: false, 
      nominalCairRiil: 0 
    };

    save({ ...safeData, keluar: [entry, ...(safeData.keluar || [])] });
    setForm({ ...emptyForm, tanggal: form.tanggal });
    flash(isSampleCategory ? "Log Sample Affiliate promosi berhasil disimpan!" : "Barang keluar dicatat");
  };

  const remove = (id) => {
    if (confirm("Apakah Anda yakin ingin menghapus data jualan/sampel ini?")) {
      save({ ...safeData, keluar: (safeData.keluar || []).filter((k) => k.id !== id) });
      flash("Data barang keluar dihapus");
    }
  };

  const months = ["SEMUA", "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"];
  const rows = (safeData.keluar || []).filter((k) => monthFilter === "SEMUA" || monthOf(k.tanggal) === monthFilter);
  const totalQty = rows.reduce((s, k) => s + (Number(k.qty) || 0), 0);

    return (
    <div>
      <SectionCard title="Catat Jualan / Barang Keluar Omnichannel" subtitle="Rekam pengeluaran stok komersial maupun alokasi biaya pengiriman sampel kreator gratis">
        
        {/* BARIS INPUT ATAS (SEJAJAR LURUS & RATA KANAN-KIRI) */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 12, marginBottom: 16 }}>
          <Field label="Tanggal Keluar">
            <input className="gm-input" type="date" value={form.tanggal} onChange={(e) => setForm({ ...form, tanggal: e.target.value })} />
          </Field>
          
          <Field label="Kode Pesanan / Resi / Nama Staf">
            <input className="gm-input" value={form.kodePesanan} onChange={(e) => setForm({ ...form, kodePesanan: e.target.value })} placeholder="Resi / nama creator promosi" />
          </Field>
          
          <Field label="Admin Shift">
            <select className="gm-select" value={form.shift} onChange={(e) => setForm({ ...form, shift: e.target.value })}>
              <option value="">Pilih Admin…</option>
              {karyawanList.map((k) => <option key={k} value={k}>{k}</option>)}
            </select>
          </Field>
          
          <Field label="Host Live">
            <select className="gm-select" value={form.host} onChange={(e) => setForm({ ...form, host: e.target.value })}>
              <option value="">Pilih Host…</option>
              {karyawanList.map((k) => <option key={k} value={k}>{k}</option>)}
            </select>
          </Field>
          
          <Field label="Channel / Jenis Keluar">
            <select className="gm-select" value={form.toko} onChange={(e) => setForm({ ...form, toko: e.target.value })}>
              {TOKO_OPTIONS.map((t) => <option key={t} value={t}>{t}</option>)}
              <option value="Sample Affiliate (Aplikasi)">🎁 Sample Affiliate (Aplikasi/Gratis Ongkir)</option>
              <option value="Sample Affiliate (Manual + Ongkir Toko)">🚚 Sample Affiliate (Manual + Ongkir Toko)</option>
            </select>
          </Field>
          
          <Field label="COD / Non COD">
            <select className="gm-select" value={form.cod} onChange={(e) => setForm({ ...form, cod: e.target.value })} disabled={isSampleCategory}>
              <option>NON COD</option>
              <option>COD</option>
            </select>
          </Field>
        </div>

        {/* BARIS INPUT BAWAH */}
        <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr 1.2fr", gap: 12, alignItems: "end" }}>
          <Field label="Pilih Produk SKU">
            <ProductSelect products={safeData.products || []} value={form.kode} onChange={(v) => setForm({ ...form, kode: v })} />
          </Field>
          <Field label="Kuantitas Qty">
            <input className="gm-input" type="number" min="1" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} />
          </Field>
          <Field label="Harga Jual Bruto Aplikasi (Rp)">
            <input 
              className="gm-input" 
              type="number" 
              value={isSampleCategory ? 0 : form.hargaJual} 
              onChange={(e) => setForm({ ...form, hargaJual: e.target.value })} 
              placeholder={isSampleCategory ? "Otomatis Rp 0" : "Harga per pcs"} 
              disabled={isSampleCategory}
            />
          </Field>
          <button className="gm-btn" onClick={submit} style={{ width: "100%", justifyContent: "center", padding: "10px 0" }}>
            <Plus size={15} /> Simpan Data Keluar
          </button>
        </div>
      </SectionCard>

      {/* TABEL DATA HISTORI JUALAN & SAMPLE */}
      <SectionCard title="Riwayat Barang Keluar & Validasi Dana Kas" subtitle={`Volume jualan terfilter: ${fmt(totalQty)} Pcs`} right={
        <select className="gm-select" style={{ width: 160 }} value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)}>
          {months.map((m) => <option key={m} value={m}>{m}</option>)}
        </select>
      }>
        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th>Tanggal</th><th>Platform / Keperluan</th><th>Invoice / Memo</th><th>SKU</th><th>Qty</th><th>Omset Bruto</th><th>Input Bersih Cair / Ongkir Sampel</th><th>Konfirmasi Status Kas</th><th>Tindakan</th>
              </tr>
            </thead>
            <tbody>
              {rows.length === 0 && <EmptyRow colSpan={9} text="Belum ada catatan barang keluar." />}
              {rows.map((k) => {
                const prod = (safeData.products || []).find((p) => p.kode === k.kode) || {};
                const isSampleApp = k.toko === "Sample Affiliate (Aplikasi)";
                const isSampleManual = k.toko === "Sample Affiliate (Manual + Ongkir Toko)";
                const isAnySample = isSampleApp || isSampleManual;

                return (
                  <tr key={k.id} style={{ backgroundColor: isAnySample ? "#fdf8f6" : "transparent" }}>
                    <td>{k.tanggal}</td>
                    <td><span style={{ color: isAnySample ? C.secondary : "inherit", fontWeight: isAnySample ? "bold" : "normal" }}>{k.toko}</span></td>
                    <td>{k.kodePesanan || "—"}</td>
                    <td style={{ fontFamily: "monospace", fontWeight: 700, color: C.primary }}>{k.kode}</td>
                    <td>{fmt(k.qty)} Pcs</td>
                    <td>{isAnySample ? <span style={{ color: C.muted }}>Sampel Gratis</span> : `Rp ${fmt(k.qty * (k.hargaJual || 0))}`}</td>
                    <td>
                      {isSampleApp && (
                        <span style={{ color: C.muted }}>HPP: Rp {fmt(k.qty * (prod.hpp || 0))} (Free Ongkir)</span>
                      )}
                      
                      {isSampleManual && (
                        <div>
                          {!k.cair ? (
                            <input 
                              type="number" 
                              className="gm-input" 
                              style={{ width: 140, padding: "4px 8px" }} 
                              placeholder="Ketik Biaya Ongkir…" 
                              value={ongkirSampleManual[k.id] || ""}
                              onChange={(e) => setOngkirSampleManual({ ...ongkirSampleManual, [k.id]: e.target.value })}
                            />
                          ) : (
                            <div style={{ fontSize: "12px" }}>
                              <div style={{ color: C.muted }}>HPP: Rp {fmt(k.qty * (prod.hpp || 0))}</div>
                              <div style={{ color: C.danger, fontWeight: "bold" }}>Ongkir: Rp {fmt(k.nominalOngkirSample || 0)}</div>
                            </div>
                          )}
                        </div>
                      )}

                      {!isAnySample && (
                        <div>
                          {!k.cair ? (
                            <input 
                              type="number" 
                              className="gm-input" 
                              style={{ width: 130, padding: "4px 8px" }} 
                              placeholder="Ketik Nominal Bersih…" 
                              value={inputCairManual[k.id] || ""}
                              onChange={(e) => setInputCairManual({ ...inputCairManual, [k.id]: e.target.value })}
                            />
                          ) : (
                            <span style={{ color: C.success, fontWeight: "bold" }}>Rp {fmt(k.nominalCairRiil)}</span>
                          )}
                        </div>
                      )}
                    </td>
                    <td>
                      <button
                        onClick={() => {
                          let nextBukuKas = [...(safeData.bukuKas || [])];

                          if (isAnySample) {
                            const totalHppSample = k.qty * (prod.hpp || 0);
                            const nominalOngkirFix = isSampleManual ? Number(ongkirSampleManual[k.id] || 0) : 0;

                            if (isSampleManual && !k.cair && !ongkirSampleManual[k.id]) {
                              alert("Mohon isi nominal biaya pengiriman ongkir yang toko bayar terlebih dahulu!");
                              return;
                            }

                            const totalBebanKredit = totalHppSample + nominalOngkirFix;
                            const nextKeluar = (safeData.keluar || []).map((x) => x.id === k.id ? { ...x, cair: !x.cair, nominalCairRiil: 0, nominalOngkirSample: nominalOngkirFix } : x);
                            
                            if (!k.cair) {
                              nextBukuKas.unshift({
                                id: uid(),
                                tanggal: todayStr(),
                                jenis: "Kredit",
                                nominal: totalBebanKredit,
                                keterangan: `[Beban Promosi Kreator] ${k.toko} SKU: ${k.kode} | Memo: ${k.kodePesanan || "Tanpa Nama"} (${k.qty} Pcs) | Terdiri dari HPP: Rp ${fmt(totalHppSample)} ${nominalOngkirFix > 0 ? `+ Ongkir Toko: Rp ${fmt(nominalOngkirFix)}` : ""}`,
                              });
                            } else {
                              nextBukuKas = nextBukuKas.filter((b) => !b.keterangan.includes(k.id) && !b.keterangan.includes(k.kodePesanan));
                            }
                            save({ ...safeData, keluar: nextKeluar, bukuKas: nextBukuKas });
                            flash("Seluruh akumulasi biaya modal & ongkir sampel sukses dibukukan sebagai Kredit!");
                          } else {
                            if (!k.cair && !inputCairManual[k.id]) {
                              alert("Mohon isi nominal pencairan bersih riil terlebih dahulu sesuai mutasi rekening/wallet toko!");
                              return;
                            }

                            const nominalCairFix = Number(inputCairManual[k.id] || 0);
                            const nextKeluar = (safeData.keluar || []).map((x) => x.id === k.id ? { ...x, cair: !x.cair, nominalCairRiil: !x.cair ? nominalCairFix : 0 } : x);

                            if (!k.cair) {
                              nextBukuKas.unshift({
                                id: uid(),
                                tanggal: todayStr(),
                                jenis: "Debit",
                                nominal: nominalCairFix,
                                keterangan: `[Pencairan Jualan Omnichannel] Toko: ${k.toko} | No Invoice: ${k.kodePesanan || "Tanpa Resi"} (${k.qty} Pcs)`,
                              });
                            } else {
                              nextBukuKas = nextBukuKas.filter((b) => !b.keterangan.includes(k.kodePesanan || k.id));
                            }
                            save({ ...safeData, keluar: nextKeluar, bukuKas: nextBukuKas });
                            flash("Mutasi nominal bersih jualan sukses masuk Buku Kas (Debit)!");
                          }
                        }}
                        style={{ fontSize: 11, fontWeight: 700, padding: "5px 12px", borderRadius: 8, border: "none", cursor: "pointer", background: k.cair ? C.success : C.primarySoft, color: k.cair ? "#fff" : C.text }}
                      >
                        {isAnySample ? (k.cair ? "✓ Terbuku di Kas" : "Bukukan Modal Sampel") : (k.cair ? "✓ Cair Masuk Kas" : "Konfirmasi Dana")}
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
  const safeData = normalizeData(data || {});
  const [form, setForm] = useState({ tanggal: todayStr(), kodePesanan: "", kode: "", kodeRetur: "", qty: "", sesuai: "Sesuai", ekspedisi: "", catatan: "" });

  const submit = () => {
    if (!form.kode || !Number(form.qty)) { flash("Pilih produk dan isi qty"); return; }
    save({ ...safeData, retur: [{ id: uid(), ...form, qty: Number(form.qty) }, ...(safeData.retur || [])] });
    setForm({ tanggal: form.tanggal, kodePesanan: "", kode: "", kodeRetur: "", qty: "", sesuai: "Sesuai", ekspedisi: "", catatan: "" });
    flash("Klaim Retur Berhadil Rekam");
  };
  const remove = (id) => save({ ...safeData, retur: (safeData.retur || []).filter((r) => r.id !== id) });

  return (
    <div>
      <SectionCard title="Manajemen Klaim Retur Masalah" subtitle="Retur berkondisi 'Sesuai' otomatis mengembalikan jumlah fisik unit ke persediaan gudang">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px,1fr))", gap: 10 }}>
          <Field label="Tanggal Klaim"><input className="gm-input" type="date" value={form.tanggal} onChange={(e) => setForm({ ...form, tanggal: e.target.value })} /></Field>
          <Field label="Nomor Invoice Awal"><input className="gm-input" value={form.kodePesanan} onChange={(e) => setForm({ ...form, kodePesanan: e.target.value })} /></Field>
          <Field label="Kode Resi Balik Retur"><input className="gm-input" value={form.kodeRetur} onChange={(e) => setForm({ ...form, kodeRetur: e.target.value })} /></Field>
          <Field label="Kurir Ekspedisi"><input className="gm-input" value={form.ekspedisi} onChange={(e) => setForm({ ...form, ekspedisi: e.target.value })} /></Field>
          <Field label="Kondisi Validasi">
            <select className="gm-select" value={form.sesuai} onChange={(e) => setForm({ ...form, sesuai: e.target.value })}>
              <option value="Sesuai">Sesuai (Balik ke Stok Gudang)</option>
              <option value="Tidak Sesuai">Tidak Sesuai (Aset Afkir / Rusak)</option>
            </select>
          </Field>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 2fr auto", gap: 10, marginTop: 10, alignItems: "end" }}>
          <Field label="Produk"><ProductSelect products={safeData.products || []} value={form.kode} onChange={(v) => setForm({ ...form, kode: v })} /></Field>
          <Field label="Qty Barang"><input className="gm-input" type="number" min="0" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} /></Field>
          <Field label="Alasan / Kronologi Kasus"><input className="gm-input" value={form.catatan} onChange={(e) => setForm({ ...form, catatan: e.target.value })} /></Field>
          <button className="gm-btn" onClick={submit}><Plus size={15} /> Simpan Retur</button>
        </div>
      </SectionCard>

      <SectionCard title="Buku Log Masalah / Retur Pakaian">
        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr><th>Tanggal</th><th>Invoice Jualan</th><th>Kode Retur</th><th>SKU Barang</th><th>Qty</th><th>Status Stok</th><th>Kurir</th><th>Keterangan</th><th></th></tr>
            </thead>
            <tbody>
              {(safeData.retur || []).length === 0 && <EmptyRow colSpan={9} text="Bersih! Tidak ada catatan retur produk." />}
              {(safeData.retur || []).map((r) => (
                <tr key={r.id}>
                  <td>{r.tanggal}</td>
                  <td>{r.kodePesanan}</td>
                  <td>{r.kodeRetur}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700 }}>{r.kode}</td>
                  <td>{r.qty} Pcs</td>
                  <td><span style={{ fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 99, background: r.sesuai === "Sesuai" ? "#E4EFE4" : "#F5E1DC", color: r.sesuai === "Sesuai" ? C.success : C.danger }}>{r.sesuai === "Sesuai" ? "Restock Gudang" : "Aset Rusak"}</span></td>
                  <td>{r.ekspedisi}</td>
                  <td>{r.catatan}</td>
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
  const safeData = normalizeData(data || {});
  const [form, setForm] = useState({ tanggal: todayStr(), kode: "", qty: "", toko: "Shopee", catatan: "" });

  const submit = () => {
    if (!form.kode || !Number(form.qty)) { flash("Pilih produk dan isi kuantitas reject!"); return; }
    save({ ...safeData, reject: [{ id: uid(), ...form, qty: Number(form.qty) }, ...(safeData.reject || [])] });
    setForm({ tanggal: todayStr(), kode: "", qty: "", toko: "Shopee", catatan: "" });
    flash("Reject dicatat memotong volume gudang");
  };
  const remove = (id) => save({ ...safeData, reject: (safeData.reject || []).filter((r) => r.id !== id) });

  return (
    <div>
      <SectionCard title="Catat Unit Reject / Cacat Jahitan" subtitle="Setiap temuan barang reject langsung memotong unit siap jualan">
        <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr 1fr 1fr auto", gap: 10, alignItems: "end" }}>
          <Field label="Tanggal Ditemukan"><input className="gm-input" type="date" value={form.tanggal} onChange={(e) => setForm({ ...form, tanggal: e.target.value })} /></Field>
          <Field label="Produk SKU"><ProductSelect products={safeData.products || []} value={form.kode} onChange={(v) => setForm({ ...form, kode: v })} /></Field>
          <Field label="Qty Rusak"><input className="gm-input" type="number" min="0" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} /></Field>
          <Field label="Lokasi Sortir">
            <select className="gm-select" value={form.toko} onChange={(e) => setForm({ ...form, toko: e.target.value })}>
              {TOKO_OPTIONS.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
          </Field>
          <button className="gm-btn" onClick={submit}><Plus size={15} /> Catat Cacat</button>
        </div>
        <div style={{ marginTop: 10 }}>
          <Field label="Detail Deskripsi Kerusakan"><input className="gm-input" value={form.catatan} onChange={(e) => setForm({ ...form, catatan: e.target.value })} placeholder="Noda noda kain / bolong kain" /></Field>
        </div>
      </SectionCard>

      <SectionCard title="Riwayat Catatan Reject Gudang">
        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th>Tanggal</th><th>Kode SKU</th><th>Nama Unit</th><th>Qty Rusak</th><th>Lokasi Sortir</th><th>Keterangan</th><th></th></tr></thead>
            <tbody>
              {(safeData.reject || []).length === 0 && <EmptyRow colSpan={7} text="Bersih! Tidak ada riwayat reject kain." />}
              {(safeData.reject || []).map((r) => {
                const prod = (safeData.products || []).find((p) => p.kode === r.kode);
                return (
                  <tr key={r.id}>
                    <td>{r.tanggal}</td>
                    <td style={{ fontFamily: "monospace", fontWeight: 700 }}>{r.kode}</td>
                    <td>{prod?.nama || "—"}</td>
                    <td style={{ color: C.danger, fontWeight: "bold" }}>{r.qty} Pcs</td>
                    <td>{r.toko}</td>
                    <td>{r.catatan}</td>
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

// ---------- Monitoring Gate Pencairan ----------
function PencairanTab({ data }) {
  const safeData = normalizeData(data || {});
  const belumCairList = (safeData.keluar || []).filter((k) => !k.cair);
  const sudahCairList = (safeData.keluar || []).filter((k) => k.cair);

  const totalEstimasiBrutoBelumCair = belumCairList.reduce((s, k) => s + (k.qty * (k.hargaJual || 0)), 0);
  const totalDanaBersihSudahCair = sudahCairList.reduce((s, k) => s + (k.nominalCairRiil || 0), 0);

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 12, marginBottom: 18 }}>
        <div className="gm-card" style={{ padding: 16, borderLeft: `5px solid ${C.danger}` }}>
          <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600 }}>DANA MENGGANTUNG APLIKASI (ESTIMASI BRUTO)</div>
          <div className="gm-h1" style={{ fontSize: 24, fontWeight: 700, marginTop: 6, color: C.danger }}>Rp {fmt(totalEstimasiBrutoBelumCair)}</div>
        </div>
        <div className="gm-card" style={{ padding: 16, borderLeft: `5px solid ${C.success}` }}>
          <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600 }}>TOTAL CAIR BERSIH RIIL MASUK KAS</div>
          <div className="gm-h1" style={{ fontSize: 24, fontWeight: 700, marginTop: 6, color: C.success }}>Rp {fmt(totalDanaBersihSudahCair)}</div>
        </div>
      </div>
    </div>
  );
}

// ---------- REKAP LAPORAN LABA RUGI TERPERINCI ----------
function LabaRugiTab({ data }) {
  const safeData = normalizeData(data || {});
  const [filterPlatform, setFilterPlatform] = useState("SEMUA");

  const barangCair = (safeData.keluar || []).filter((k) => k.cair);
  const filteredKeluar = barangCair.filter((k) => filterPlatform === "SEMUA" || k.toko === filterPlatform);

  const totalOmsetBersihCair = filteredKeluar.reduce((s, k) => s + (Number(k.nominalCairRiil) || 0), 0);
  const totalHppModalTerjual = filteredKeluar.reduce((s, k) => {
    const prod = (safeData.products || []).find((p) => p.kode === k.kode) || {};
    return s + ((Number(k.qty) || 0) * (prod.hpp || 0));
  }, 0);
  const keuntunganNettoUsaha = totalOmsetBersihCair - totalHppModalTerjual;

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12, marginBottom: 20 }}>
        <div className="gm-card" style={{ padding: 18, borderLeft: `5px solid ${C.success}` }}>
          <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600 }}>PENDAPATAN CAIR BERSIH (OMSET RIIL)</div>
          <div className="gm-h1" style={{ fontSize: 24, fontWeight: 700, color: C.success, marginTop: 6 }}>Rp {fmt(totalOmsetBersihCair)}</div>
        </div>
        <div className="gm-card" style={{ padding: 18, borderLeft: `5px solid ${C.secondary}` }}>
          <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600 }}>TOTAL BEBAN POKOK (HPP MODAL BARANG)</div>
          <div className="gm-h1" style={{ fontSize: 24, fontWeight: 700, color: C.secondary, marginTop: 6 }}>Rp {fmt(totalHppModalTerjual)}</div>
        </div>
        <div className="gm-card" style={{ padding: 18, borderLeft: `5px solid ${keuntunganNettoUsaha >= 0 ? C.accent : C.danger}` }}>
          <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600 }}>LABA RUGI BERSIH OPERASIONAL</div>
          <div className="gm-h1" style={{ fontSize: 24, fontWeight: 700, color: keuntunganNettoUsaha >= 0 ? C.accent : C.danger, marginTop: 6 }}>Rp {fmt(keuntunganNettoUsaha)}</div>
        </div>
      </div>

      <SectionCard 
        title="Laporan Laba Rugi Komparasi Platform Omnichannel" 
        subtitle="Analisis keuntungan bersih riil berdasarkan omset pencairan dana bersih dikurangi beban pokok HPP baju terjual"
        right={
          <select className="gm-select" style={{ width: 160 }} value={filterPlatform} onChange={(e) => setFilterPlatform(e.target.value)}>
            <option value="SEMUA">Semua Platform</option>
            {TOKO_OPTIONS.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>
        }
      >
        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th>Tanggal</th><th>Platform Toko</th><th>No. Invoice Jualan</th><th>Kode SKU</th><th>Nama Pakaian</th><th>Volume</th><th>Pencairan Bersih</th><th>Beban HPP Modal</th><th>Laba Bersih Riil</th>
              </tr>
            </thead>
            <tbody>
              {filteredKeluar.length === 0 && <EmptyRow colSpan={9} text="Belum ada data penjualan cair untuk platform terpilih." />}
              {filteredKeluar.map((k) => {
                const prod = (safeData.products || []).find((p) => p.kode === k.kode) || {};
                const modalHppRow = (Number(k.qty) || 0) * (prod.hpp || 0);
                const labaRow = (Number(k.nominalCairRiil) || 0) - modalHppRow;
                return (
                  <tr key={k.id}>
                    <td>{k.tanggal}</td>
                    <td><strong>{k.toko}</strong></td>
                    <td>{k.kodePesanan || "—"}</td>
                    <td style={{ fontFamily: "monospace", fontWeight: 700 }}>{k.kode}</td>
                    <td style={{ whiteSpace: "normal" }}>{prod.nama || "—"}</td>
                    <td>{k.qty} Pcs</td>
                    <td style={{ color: C.success, fontWeight: 600 }}>Rp {fmt(k.nominalCairRiil)}</td>
                    <td style={{ color: C.secondary }}>Rp {fmt(modalHppRow)}</td>
                    <td style={{ color: labaRow >= 0 ? C.accent : C.danger, fontWeight: 700 }}>Rp {fmt(labaRow)}</td>
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

// ---------- Master Buku Kas ----------
function BukuKasTab({ data, save, flash }) {
  const safeData = normalizeData(data || {});
  const [formJurnal, setFormJurnal] = useState({ tanggal: todayStr(), jenis: "Debit", nominal: "", keterangan: "" });

  const handleJurnalManual = () => {
    if (!Number(formJurnal.nominal)) { flash("Masukkan nominal uang!"); return; }
    const entry = { id: uid(), ...formJurnal, nominal: Number(formJurnal.nominal) };
    save({ ...safeData, bukuKas: [entry, ...(safeData.bukuKas || [])] });
    setFormJurnal({ tanggal: todayStr(), jenis: "Debit", nominal: "", keterangan: "" });
    flash("Jurnal keuangan disimpan!");
  };

  const removeJurnal = (id) => {
    if (confirm("Hapus catatan jurnal kas ini?")) {
      save({ ...safeData, bukuKas: (safeData.bukuKas || []).filter((b) => b.id !== id) });
      flash("Jurnal berhasil dihapus");
    }
  };

  const totalDebit = (safeData.bukuKas || []).filter(b => b.jenis === "Debit").reduce((s, b) => s + Number(b.nominal || 0), 0);
  const totalKredit = (safeData.bukuKas || []).filter(b => b.jenis === "Kredit").reduce((s, b) => s + Number(b.nominal || 0), 0);
  const saldoNetto = totalDebit - totalKredit;

  const sortedKronologis = [...(safeData.bukuKas || [])].sort((a, b) => new Date(a.tanggal) - new Date(b.tanggal));
  let runningSaldo = 0;
  const kasWithSaldo = sortedKronologis.map(b => {
    if (b.jenis === "Debit") runningSaldo += b.nominal;
    else runningSaldo -= b.nominal;
    return { ...b, saldoKumulatif: runningSaldo };
  }).reverse();

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 12, marginBottom: 20 }}>
        <div className="gm-card" style={{ padding: 16 }}><small style={{ color: C.muted }}>TOTAL DANA CAIR MASUK (DEBIT)</small><h3 style={{ color: C.success }}>Rp {fmt(totalDebit)}</h3></div>
        <div className="gm-card" style={{ padding: 16 }}><small style={{ color: C.muted }}>TOTAL BIAYA KELUAR (KREDIT)</small><h3 style={{ color: C.danger }}>Rp {fmt(totalKredit)}</h3></div>
        <div className="gm-card" style={{ padding: 16 }}><small style={{ color: C.muted }}>SALDO NETTO KAS AKTIF TOKO</small><h3 style={{ color: C.primary }}>Rp {fmt(saldoNetto)}</h3></div>
      </div>

      <SectionCard title="Input Manual Jurnal Buku Kas">
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 2fr auto", gap: 10, alignItems: "end" }}>
          <Field label="Tanggal"><input className="gm-input" type="date" value={formJurnal.tanggal} onChange={(e) => setFormJurnal({ ...formJurnal, tanggal: e.target.value })} /></Field>
          <Field label="Jenis Arus Kas">
            <select className="gm-select" value={formJurnal.jenis} onChange={(e) => setFormJurnal({ ...formJurnal, jenis: e.target.value })}>
              <option value="Debit">Debit (Uang Masuk / Modal)</option>
              <option value="Kredit">Kredit (Uang Keluar / Biaya)</option>
            </select>
          </Field>
          <Field label="Nominal Uang (Rp)"><input className="gm-input" type="number" value={formJurnal.nominal} onChange={(e) => setFormJurnal({ ...formJurnal, nominal: e.target.value })} /></Field>
          <Field label="Keterangan Nota"><input className="gm-input" value={formJurnal.keterangan} onChange={(e) => setFormJurnal({ ...formJurnal, keterangan: e.target.value })} /></Field>
          <button className="gm-btn" onClick={handleJurnalManual}>Catat Kas</button>
        </div>
      </SectionCard>

      <SectionCard title="Rekening Koran & Saldo Berjalan Kronologis">
        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th>No</th><th>Tanggal</th><th>Deskripsi Transaksi</th><th>Debit (+)</th><th>Kredit (-)</th><th>Saldo Kumulatif</th><th>Aksi</th></tr></thead>
            <tbody>
              {kasWithSaldo.length === 0 && <EmptyRow colSpan={7} text="Belum ada transaksi kas." />}
              {kasWithSaldo.map((b, idx) => (
                <tr key={b.id}>
                  <td>{idx + 1}</td><td>{b.tanggal}</td><td>{b.keterangan}</td>
                  <td style={{ color: C.success }}>{b.jenis === "Debit" ? `Rp ${fmt(b.nominal)}` : "—"}</td>
                  <td style={{ color: C.danger }}>{b.jenis === "Kredit" ? `Rp ${fmt(b.nominal)}` : "—"}</td>
                  <td style={{ fontWeight: "bold" }}>Rp {fmt(b.saldoKumulatif)}</td>
                  <td><DeleteBtn onClick={() => removeJurnal(b.id)} /></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </SectionCard>
    </div>
  );
}

// ---------- Keuangan & Gaji Tim ----------
function KeuanganTab({ data, save, flash }) {
  const safeData = normalizeData(data || {});
  const [panel, setPanel] = useState("gaji-dinamis");
  const [formKasbon, setFormKasbon] = useState({ tanggal: todayStr(), nama: "", nominal: "", keterangan: "" });
  const [gajiDinamis, setGajiDinamis] = useState(safeData.gajiDinamis || {});

  useEffect(() => {
    if (safeData.gajiDinamis && Object.keys(safeData.gajiDinamis).length > 0) {
      setGajiDinamis(safeData.gajiDinamis);
    }
  }, [safeData.gajiDinamis]);

  const karyawanList = (safeData.karyawanMaster || []).map((k) => k.nama).filter(Boolean) || [];

  const submitKasbon = () => {
    if (!formKasbon.nama.trim() || !Number(formKasbon.nominal)) { flash("Isi nama karyawan and nominal kasbon!"); return; }
    const entry = { id: uid(), tanggal: formKasbon.tanggal, nama: formKasbon.nama, nominal: Number(formKasbon.nominal), keterangan: formKasbon.keterangan };
    
    save({
      ...safeData,
      kasbon: [entry, ...(safeData.kasbon || [])],
      bukuKas: [{ id: uid(), tanggal: formKasbon.tanggal, jenis: "Kredit", nominal: Number(formKasbon.nominal), keterangan: `[Kasbon Pegawai] Nama: ${formKasbon.nama} | Memo: ${formKasbon.keterangan || "Pinjaman"}` }, ...(safeData.bukuKas || [])]
    });
    setFormKasbon({ tanggal: todayStr(), nama: "", nominal: "", keterangan: "" });
    flash("Kasbon pegawai dicatat memotong Buku Kas utama");
  };

  const saveSalarySettings = () => {
    save({ ...safeData, gajiDinamis: gajiDinamis });
    flash("Konfigurasi parameter nilai gaji manual berhasil disimpan!");
  };

  const rekapSlips = useMemo(() => {
    const slips = [];
    const empMap = {};

    (safeData.karyawanMaster || []).forEach((e) => {
      empMap[e.nama] = { nama: e.nama, role: e.peran || "Tim Live", qtyKeluar: 0 };
    });

    (safeData.keluar || []).forEach((k) => {
      if (k.host && empMap[k.host]) empMap[k.host].qtyKeluar += Number(k.qty) || 0;
      if (k.shift && k.shift !== k.host && empMap[k.shift]) empMap[k.shift].qtyKeluar += Number(k.qty) || 0;
    });

    const kasbonMap = {};
    (safeData.kasbon || []).forEach((kb) => {
      kasbonMap[kb.nama] = (kasbonMap[kb.nama] || 0) + (Number(kb.nominal) || 0);
    });

    Object.values(empMap).forEach((emp) => {
      const gPokok = Number(gajiDinamis[emp.nama]?.gajiPokok) || 0;
      const kPcs = Number(gajiDinamis[emp.nama]?.komisiPerPcs) || 0;
      const bruto = gPokok + (emp.qtyKeluar * kPcs);
      const potonganKb = kasbonMap[emp.nama] || 0;
      
      slips.push({ ...emp, bruto, kasbon: potonganKb, bersih: Math.max(0, bruto - potonganKb) });
    });
    return slips;
  }, [safeData.karyawanMaster, safeData.keluar, safeData.kasbon, gajiDinamis]);

  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
        <button className="gm-btn" style={{ background: panel === "gaji-dinamis" ? C.primary : C.border, color: panel === "gaji-dinamis" ? "#fff" : C.text }} onClick={() => setPanel("gaji-dinamis")}>⚙️ Atur Gaji Manual</button>
        <button className="gm-btn" style={{ background: panel === "kasbon" ? C.primary : C.border, color: panel === "kasbon" ? "#fff" : C.text }} onClick={() => setPanel("kasbon")}>💰 Ambil Kasbon</button>
        <button className="gm-btn" style={{ background: panel === "slip" ? C.primary : C.border, color: panel === "slip" ? "#fff" : C.text }} onClick={() => setPanel("slip")}>📋 Slip Gaji Tim</button>
      </div>

      {panel === "gaji-dinamis" && (
        <SectionCard title="Pengaturan Gaji Pokok & Komisi Insentif" subtitle="Tentukan nilai gaji tetap dan besaran bonus per pcs terjual untuk setiap staf operasional">
          {karyawanList.length === 0 ? (
            <div style={{ padding: 15, color: C.muted, fontSize: 13 }}>Silakan tambahkan data karyawan di menu tab Kode Master terlebih dahulu.</div>
          ) : (
            <div>
              <table className="gm-table" style={{ width: "100%", marginBottom: 15 }}>
                <thead><tr><th>Nama Pegawai</th><th>Gaji Pokok Tetap (Rp)</th><th>Insentif Komisi per Pcs (Rp)</th></tr></thead>
                <tbody>
                  {karyawanList.map((nama) => (
                    <tr key={nama}>
                      <td style={{ fontWeight: 600 }}>{nama}</td>
                      <td><input className="gm-input" type="number" value={gajiDinamis[nama]?.gajiPokok || ""} onChange={(e) => setGajiDinamis({ ...gajiDinamis, [nama]: { ...(gajiDinamis[nama] || {}), gajiPokok: e.target.value } })} placeholder="Rp 0" /></td>
                      <td><input className="gm-input" type="number" value={gajiDinamis[nama]?.komisiPerPcs || ""} onChange={(e) => setGajiDinamis({ ...gajiDinamis, [nama]: { ...(gajiDinamis[nama] || {}), komisiPerPcs: e.target.value } })} placeholder="Rp 0" /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <button className="gm-btn" onClick={saveSalarySettings}><Save size={14} /> Simpan Struktur Gaji</button>
            </div>
          )}
        </SectionCard>
      )}

      {panel === "kasbon" && (
        <div>
          <SectionCard title="Pencatatan Ambil Dana Pinjaman Kasbon Karyawan">
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1.5fr 1fr 2fr auto", gap: 10, alignItems: "end" }}>
              <Field label="Tanggal"><input className="gm-input" type="date" value={formKasbon.tanggal} onChange={(e) => setFormKasbon({ ...formKasbon, tanggal: e.target.value })} /></Field>
              <Field label="Nama Pegawai">
                <select className="gm-select" value={formKasbon.nama} onChange={(e) => setFormKasbon({ ...formKasbon, nama: e.target.value })}>
                  <option value="">Pilih Pegawai…</option>
                  {karyawanList.map((k) => <option key={k} value={k}>{k}</option>)}
                </select>
              </Field>
              <Field label="Nominal Pinjaman"><input className="gm-input" type="number" value={formKasbon.nominal} onChange={(e) => setFormKasbon({ ...formKasbon, nominal: e.target.value })} /></Field>
              <Field label="Memo / Alasan"><input className="gm-input" value={formKasbon.keterangan} onChange={(e) => setFormKasbon({ ...formKasbon, keterangan: e.target.value })} placeholder="Bensin live / makan" /></Field>
              <button className="gm-btn" onClick={submitKasbon}>Simpan Kasbon</button>
            </div>
          </SectionCard>

          <SectionCard title="Histori Log Pinjaman Kasbon">
            <div style={{ overflowX: "auto" }}>
              <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead><tr><th>Tanggal</th><th>Nama Pegawai</th><th>Nominal Pinjaman</th><th>Keterangan</th></tr></thead>
                <tbody>
                  {(safeData.kasbon || []).length === 0 && <EmptyRow colSpan={4} text="Belum ada catatan kasbon pegawai." />}
                  {(safeData.kasbon || []).map((k) => (
                    <tr key={k.id}>
                      <td>{k.tanggal}</td>
                      <td><strong>{k.nama}</strong></td>
                      <td style={{ color: C.danger, fontWeight: "bold" }}>Rp {fmt(k.nominal)}</td>
                      <td>{k.keterangan}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </SectionCard>
        </div>
      )}

      {panel === "slip" && (
        <SectionCard title="Rekap Slip Gaji Karyawan Otomatis" subtitle="Kalkulasi upah bruto dihitung terintegrasi dari gabungan akumulasi volume jualan live/shift tim">
          <div style={{ overflowX: "auto" }}>
            <table className="gm-table" style={{ width: "100%" }}>
              <thead><tr><th>Nama Pegawai</th><th>Tugas Divisi</th><th>Total Jualan</th><th>Gaji Bruto</th><th>Potongan Kasbon</th><th>Sisa Gaji Bersih</th></tr></thead>
              <tbody>
                {rekapSlips.map((s) => (
                  <tr key={s.nama}>
                    <td style={{ fontWeight: 600 }}>{s.nama}</td>
                    <td>{s.role}</td>
                    <td>{fmt(s.qtyKeluar)} Pcs</td>
                    <td style={{ color: C.primary, fontWeight: 600 }}>Rp {fmt(s.bruto)}</td>
                    <td style={{ color: C.danger }}>Rp {fmt(s.kasbon)}</td>
                    <td style={{ color: C.success, fontWeight: 700 }}>Rp {fmt(s.bersih)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </SectionCard>
      )}
    </div>
  );
}

// ---------- Master Kode Variable ----------
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
    const newEmployee = { id: uid(), nama: formKaryawan.nama.toUpperCase().trim(), peran: formKaryawan.peran };
    save({ ...safeData, karyawanMaster: [...karyawanList, newEmployee] });
    setFormKaryawan({ nama: "", peran: "Host Live" });
    flash("Staf operational didaftarkan");
  };

  const addMaster = () => {
    if (!form.code.trim() || !form.name.trim()) { flash("Isi inisial kode dan arti nama!"); return; }
    const next = { ...safeData, masterCode: { ...safeData.masterCode, [cat]: [...listData, { code: form.code.trim().toUpperCase(), name: form.name.trim().toUpperCase() }] } };
    save(next);
    setForm({ code: "", name: "" });
    flash("Kode master baru disimpan");
  };

  return (
    <div>
      <SectionCard title="Setup Kamus Referensi Kode Master">
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 16, overflowX: "auto" }}>
          {KATEGORI.map((k) => (
            <button key={k.key} className="gm-btn" style={{ background: cat === k.key ? C.primary : C.surfaceAlt, color: cat === k.key ? "#fff" : C.text }} onClick={() => setCat(k.key)}>{k.label}</button>
          ))}
        </div>
        {cat === "karyawanMaster" ? (
          <div style={{ display: "grid", gridTemplateColumns: "2fr 1.5fr auto", gap: 10, alignItems: "end" }}>
            <Field label="Nama Lengkap Pegawai"><input className="gm-input" value={formKaryawan.nama} onChange={(e) => setFormKaryawan({ ...formKaryawan, nama: e.target.value })} /></Field>
            <Field label="Jabatan"><select className="gm-select" value={formKaryawan.peran} onChange={(e) => setFormKaryawan({ ...formKaryawan, peran: e.target.value })}><option value="Host Live">Host Live Streaming</option><option value="Admin Live">Admin Live / Shift</option></select></Field>
            <button className="gm-btn" onClick={addKaryawan}>Registrasi Tim</button>
          </div>
        ) : (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr auto", gap: 10, alignItems: "end" }}>
            <Field label="Inisial Kode SKU"><input className="gm-input" value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} /></Field>
            <Field label="Arti Deskripsi"><input className="gm-input" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
            <button className="gm-btn" onClick={addMaster}>Simpan Kode</button>
          </div>
        )}
      </SectionCard>
    </div>
  );
}

// ---------- Render Laporan Cetak PDF ----------
function PrintTable({ title, rows }) {
  if (!rows || !rows.length) return null;
  const headers = Object.keys(rows);
  return (
    <div style={{ marginBottom: 22, breakInside: "avoid" }}>
      <h3 style={{ fontFamily: "Georgia, serif", color: C.primaryDark, fontSize: 14 }}>{title}</h3>
      <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "10px" }} border="1" cellPadding="5" borderColor="#ccc">
        <thead><tr style={{ background: "#F3ECE1" }}>{headers.map((h) => <th key={h} style={{ textAlign: "left" }}>{h}</th>)}</tr></thead>
        <tbody>{rows.map((r, i) => <tr key={i}>{headers.map((h) => <td key={h}>{String(r[h] ?? "")}</td>)}</tr>)}</tbody>
      </table>
    </div>
  );
}

function PrintReport({ data, stockMap, stokAkhir, onClose }) {
  const { produk, masuk, keluar, retur, reject } = buildRows(data, stockMap, stokAkhir);
  useEffect(() => { setTimeout(() => window.print(), 500); }, []);
  return (
    <div style={{ padding: 30, background: "#fff" }}>
      <div className="no-print" style={{ marginBottom: 20 }}><button className="gm-btn" onClick={onClose}>← Kembali</button></div>
      <h2>Laporan Cetak Fisik Pembukuan Niki Dziyab</h2>
      <PrintTable title="1. Ringkasan Aset Saldo Unit Persediaan Gudang" rows={produk} />
      <PrintTable title="2. Dokumen Kelola Mutasi Barang Masuk" rows={masuk} />
      <PrintTable title="3. Laporan Buku Kas Keluar (Transaksi Penjualan)" rows={keluar} />
      <PrintTable title="4. Rekapitulasi Klaim Retur Masalah Pembeli" rows={retur} />
      <PrintTable title="5. Buku Catatan Afkir / Reject Produksi Kain" rows={reject} />
      <style>{`@media print { .no-print { display: none !important; } }`}</style>
    </div>
  );
}
