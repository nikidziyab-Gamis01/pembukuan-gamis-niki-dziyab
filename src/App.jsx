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

const TOKO_OPTIONS = ["Shopee", "TikTok", "Tokopedia", "Lazada", "Offline"];
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
    gajiSetelan: {
      gajiPokokOwner: Number(source.gajiSetelan?.gajiPokokOwner) || 0,
      komisiOwnerPerPcs: Number(source.gajiSetelan?.komisiOwnerPerPcs) || 0,
      komisiHostPerPcs: Number(source.gajiSetelan?.komisiHostPerPcs) || 0,
      komisiAdminPerPcs: Number(source.gajiSetelan?.komisiAdminPerPcs) || 0,
    },
  };
};

const EMPTY_DATA = normalizeData({
  products: [],
  masuk: [],
  keluar: [],
  retur: [],
  reject: [],
  masterCode: EMPTY_MASTER,
  kasbon: [],
  karyawanMaster: [],
  bukuKas: [],
  gajiSetelan: { gajiPokokOwner: 0, komisiOwnerPerPcs: 0, komisiHostPerPcs: 0, komisiAdminPerPcs: 0 },
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
    kasbon: [],
    gajiSetelan: { gajiPokokOwner: 0, komisiOwnerPerPcs: 0, komisiHostPerPcs: 0, komisiAdminPerPcs: 0 },
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
  { id: "bukuKas", label: "📊 Master Buku Kas", icon: PieChart },
  { id: "keuangan", label: "💰 Keuangan & Gaji Tim", icon: Settings2 },
  { id: "master", label: "⚙️ Kode Master", icon: Settings2 },
];

export default function GamisInventoryApp() {
    const [data, setData] = useState(EMPTY_DATA);

  const [ready, setReady] = useState(false);
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
          // Pengaman mutlak: Email utama owner pertama langsung lolos tanpa antre database
          if (currentUser.email.toLowerCase() === "nikidziyab@gmail.com") {
            setUser(currentUser);
            setAuthLoading(false);
            return;
          }

          // Mengambil daftar email tambahan secara langsung dari Firestore
          const whitelistRef = doc(db, "pembukuan_gamis", "whitelist_users");
          const whitelistSnap = await getDoc(whitelistRef);
          
          let allowedEmails = [];
          if (whitelistSnap.exists()) {
            const allowedData = whitelistSnap.data();
            if (allowedData && Array.isArray(allowedData.emails)) {
              allowedEmails = allowedData.emails.map((e) => String(e || "").toLowerCase());
            }
          }

          // Validasi email tambahan dari Firestore
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
      } else if (code === "auth/configuration-not-found") {
        setAuthError("Firebase Auth belum dikonfigurasi dengan benar. Periksa file .env dan domain otorisasi di Firebase Console.");
      } else {
        setAuthError("Login Google gagal. Cek konfigurasi Firebase dan izin browser.");
      }

      alert(
        code === "auth/popup-closed-by-user"
          ? "Login dibatalkan. Silakan klik tombol Google kembali."
          : code === "auth/popup-blocked"
            ? "Popup diblokir browser. Izinkan popup untuk aplikasi ini lalu coba lagi."
            : "Login Google gagal. Cek konfigurasi Firebase dan izin browser."
      );
    }
  };

  const logout = async () => {
    try {
      await signOut(auth);
    } catch (error) {
      console.error("Gagal melakukan sign out:", error);
    }
  };

      useEffect(() => {
    // Membaca data dari Firestore secara realtime dari dokumen "data_utama"
    const unsubscribe = onSnapshot(doc(db, "pembukuan_gamis", "data_utama"), (snapshot) => {
      if (snapshot.exists()) {
        const cloudData = snapshot.data();
        const payload = cloudData && cloudData.pembukuanData ? cloudData.pembukuanData : EMPTY_DATA;
        setData(normalizeData(payload));
      } else {
        setData(EMPTY_DATA);
      }
    });

    return () => unsubscribe();
  }, []);

  const syncToCloud = async (nextData = data) => {
    const normalized = normalizeData(nextData);
    setData(normalized);

    try {
      const docRef = doc(db, "pembukuan_gamis", "data_utama");
      await setDoc(docRef, {
        pembukuanData: normalized,
        updatedAt: new Date()
      });

      console.log("Data berhasil sinkron otomatis ke Cloud Firebase!");
      return true;
    } catch (e) {
      console.error("Gagal menyimpan data ke Firebase:", e);
      return false;
    }
  };

  const save = async (next) => {
    const normalized = normalizeData(next);
    setData(normalized);
    return syncToCloud(normalized);
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

    const saveAsFile = () => {
    try {
      const payload = JSON.stringify(normalizeData(data), null, 2);
      downloadBlob(payload, `pembukuan-gamis-backup-${todayStr()}.json`, "application/json");
      flash("File backup berhasil diunduh");
    } catch (error) {
      console.error("Gagal mengunduh file cadangan:", error);
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
          <button onClick={loginWithGoogle} style={{ backgroundColor: "#10b981", color: "#fff", border: "none", padding: "12px 24px", borderRadius: 8, fontSize: 15, fontWeight: "bold", cursor: "pointer", display: "flex", alignItems: "center", gap: 10, margin: "0 auto", transition: "background-color 0.2s" }}>
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
                <div style={{ fontSize: 11.5, color: C.muted }}>Stok · Barang Masuk & Keluar · Retur · Reject · Pencairan · Keuangan</div>
              </div>
            </div>

            <div style={{ display: "flex", alignItems: "center", gap: 6, position: "relative", flexWrap: "wrap", justifyContent: "flex-end" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, background: C.surfaceAlt, border: `1px solid ${C.border}`, borderRadius: 999, padding: "6px 10px 6px 8px", marginRight: 4 }}>
                {user?.photoURL ? (
                  <img src={user.photoURL} alt="Avatar user" style={{ width: 26, height: 26, borderRadius: "50%", objectFit: "cover", border: `1px solid ${C.border}` }} />
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

              <button className="gm-btn-ghost" style={{ borderRadius: 8, padding: "8px 12px", fontSize: 12.5, fontWeight: 600, display: "inline-flex", alignItems: "center", gap: 6, cursor: "pointer", background: "transparent" }} onClick={logout} title="Keluar dari akun Google saat ini">
                <XCircle size={14} /> Logout
              </button>

              <input ref={fileInputRef} type="file" accept="application/json" onChange={handleImportFile} style={{ display: "none" }} />
              <button className="gm-btn-ghost" style={{ borderRadius: 8, padding: "8px 12px", fontSize: 12.5, fontWeight: 600, display: "inline-flex", alignItems: "center", gap: 6, cursor: "pointer", background: "transparent" }} onClick={() => fileInputRef.current?.click()} title="Muat file cadangan (.json)">
                <Upload size={14} /> Impor
              </button>
              <button className="gm-btn-ghost" style={{ borderRadius: 8, padding: "8px 12px", fontSize: 12.5, fontWeight: 600, display: "inline-flex", alignItems: "center", gap: 6, cursor: "pointer", background: "transparent" }} onClick={saveAsFile} title="Download file backup data .json untuk cadangan lokal">
                <Download size={14} /> Download File Backup
              </button>
              <button className="gm-btn" style={{ padding: "8px 14px", fontSize: 12.5, background: C.success }} onClick={syncCloudNow} title="Simpan/update data ke Firebase secara realtime untuk tim">
                <Save size={14} /> Update Real-time
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
        {tab === "bukuKas" && <BukuKasTab data={data} save={save} flash={flash} />}
        {tab === "keuangan" && <KeuanganTab data={data} save={save} flash={flash} />}
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
  const safeData = normalizeData(data || {});
  const totalProduk = (safeData?.products?.length) || 0;
  const totalStok = (safeData?.products || []).reduce((s, p) => s + stokAkhir(p?.kode), 0);
  const bulanIni = monthOf(todayStr());
  const keluarBulanIni = (safeData?.keluar || []).filter((k) => monthOf(k?.tanggal) === bulanIni).reduce((s, k) => s + (Number(k?.qty) || 0), 0);
  const returBulanIni = (safeData?.retur || []).filter((r) => monthOf(r?.tanggal) === bulanIni).reduce((s, r) => s + (Number(r?.qty) || 0), 0);
  const rejectBulanIni = (safeData?.reject || []).filter((r) => monthOf(r?.tanggal) === bulanIni).reduce((s, r) => s + (Number(r?.qty) || 0), 0);

  const totalKeluar = (safeData?.keluar || []).reduce((s, k) => s + (Number(k?.qty) || 0), 0);
  const totalCair = (safeData?.keluar || []).filter((k) => k?.cair).reduce((s, k) => s + (Number(k?.qty) || 0), 0);

  // Hitung Keuntungan Bersih Toko dari bukuKas (Debit - Kredit)
  const totalDebit = (safeData?.bukuKas || []).filter((b) => b?.jenis === "Debit").reduce((s, b) => s + (Number(b?.nominal) || 0), 0);
  const totalKredit = (safeData?.bukuKas || []).filter((b) => b?.jenis === "Kredit").reduce((s, b) => s + (Number(b?.nominal) || 0), 0);
  const keuntunganBersih = totalDebit - totalKredit;

  // Hitung total beban gaji dari karyawan
  const totalBebanGaji = calculateTotalSalaryLoad(safeData);

  const chartData = BULAN_ID.slice(6, 12).map((b) => ({
    bulan: b.slice(0, 3),
    qty: (safeData?.keluar || []).filter((k) => monthOf(k?.tanggal) === b).reduce((s, k) => s + (Number(k?.qty) || 0), 0),
  }));

  const lowStock = (safeData?.products || []).filter((p) => stokAkhir(p?.kode) <= 10);

  const cards = [
    { label: "Total Produk", value: fmt(totalProduk), color: C.primary },
    { label: "Stok Tersedia", value: fmt(totalStok), color: C.success },
    { label: `Keluar — ${bulanIni}`, value: fmt(keluarBulanIni), color: C.accent },
    { label: `Retur — ${bulanIni}`, value: fmt(returBulanIni), color: C.danger },
    { label: `Reject — ${bulanIni}`, value: fmt(rejectBulanIni), color: C.muted },
    { label: "Keuntungan Bersih", value: fmt(keuntunganBersih), color: keuntunganBersih >= 0 ? C.success : C.danger },
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
              <BarChart data={chartData || []} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={C.border} vertical={false} />
                <XAxis dataKey="bulan" tick={{ fontSize: 12, fill: C.muted }} axisLine={{ stroke: C.border }} tickLine={false} />
                <YAxis tick={{ fontSize: 12, fill: C.muted }} axisLine={false} tickLine={false} />
                <Tooltip contentStyle={{ borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 12 }} />
                <Bar dataKey="qty" radius={[5, 5, 0, 0]}>
                  {(chartData || []).map((entry, i) => (
                    <Cell key={i} fill={entry?.bulan === bulanIni?.slice(0, 3) ? C.primary : C.accentSoft} />
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

      <SectionCard title="Stok Menipis" subtitle="Produk dengan stok akhir ≤ 10" right={lowStock?.length > 0 && <AlertTriangle size={16} color={C.danger} />}>
        {lowStock?.length === 0 ? (
          <div style={{ fontSize: 13, color: C.muted }}>Tidak ada produk dengan stok menipis.</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {(lowStock || []).map((p) => (
              <div key={p?.kode} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 10px", background: C.surfaceAlt, borderRadius: 8 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ width: 10, height: 10, borderRadius: "50%", background: WARNA_SWATCH[p?.warna] || C.muted, display: "inline-block" }} />
                  <span style={{ fontSize: 13, fontWeight: 500 }}>{p?.kode} — {p?.nama}</span>
                </div>
                <span style={{ fontSize: 13, fontWeight: 700, color: C.danger }}>{stokAkhir(p?.kode)}</span>
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
  const safeData = normalizeData(data || {});
  const products = (safeData?.products || []) || [];
  const mc = (safeData?.masterCode || {}) || { vendor: [], jenis: [], brand: [], bahan: [], warna: [], ukuran: [] };
  const [form, setForm] = useState({ vendor: "", jenis: "", brand: "", bahan: "", warna: "", ukuran: "", nama: "", stokAwal: "" });

  const codeOf = (list, name) => (Array.isArray(list) ? (list.find((x) => x?.name === name) || {}).code : "") || "";
  const previewKode =
    form.vendor && form.jenis && form.brand && form.bahan && form.warna && form.ukuran
      ? codeOf(mc?.vendor, form.vendor) + codeOf(mc?.jenis, form.jenis) + codeOf(mc?.brand, form.brand) + codeOf(mc?.bahan, form.bahan) + codeOf(mc?.warna, form.warna) + codeOf(mc?.ukuran, form.ukuran)
      : "";

  const addProduk = () => {
    if (!previewKode) { flash("Lengkapi vendor, jenis, brand, bahan, warna, dan ukuran dulu"); return; }
    if ((products || []).some((p) => p?.kode === previewKode)) { flash("Kode barang sudah ada"); return; }
    const nama = form.nama.trim() || `${form.brand} ${form.jenis} ${form.warna} ${form.ukuran} ${form.vendor}`.toUpperCase();
    const next = { ...safeData, products: [...(products || []), { kode: previewKode, nama, jenis: form.jenis, brand: form.brand, bahan: form.bahan, warna: form.warna, ukuran: form.ukuran, vendor: form.vendor }] };
    if (Number(form.stokAwal) > 0) {
      next.masuk = [...(safeData?.masuk || []), { id: uid(), tanggal: todayStr(), kode: previewKode, qty: Number(form.stokAwal), vendor: form.vendor, catatan: "Stok awal" }];
    }
    save(next);
    setForm({ vendor: "", jenis: "", brand: "", bahan: "", warna: "", ukuran: "", nama: "", stokAwal: "" });
    flash(`Produk ${previewKode} ditambahkan`);
  };

  const removeProduk = (kode) => {
    const hasLogs = [...(safeData?.masuk || []), ...(safeData?.keluar || []), ...(safeData?.retur || []), ...(safeData?.reject || [])].some((r) => r?.kode === kode);
    if (hasLogs) { flash("Tidak bisa dihapus — produk ini punya riwayat transaksi"); return; }
    save({ ...safeData, products: (products || []).filter((p) => p?.kode !== kode) });
  };

  const filtered = (products || []).filter((p) => {
    if (!p) return false;
    const kataKunci = q.trim().toLowerCase();
    if (!kataKunci) return true;
    const kode = String(p?.kode || "").toLowerCase();
    const nama = String(p?.nama || "").toLowerCase();
    const warna = String(p?.warna || "").toLowerCase();
    const brand = String(p?.brand || "").toLowerCase();
    return kode.includes(kataKunci) || nama.includes(kataKunci) || warna.includes(kataKunci) || brand.includes(kataKunci);
  });

  return (
    <div>
      <SectionCard title="Tambah Produk" subtitle="Kode dibuat otomatis dari Kode Master (Vendor + Jenis + Brand + Bahan + Warna + Ukuran)">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px,1fr))", gap: 10 }}>
          <Field label="Vendor">
            <select className="gm-select" value={form.vendor} onChange={(e) => setForm({ ...form, vendor: e.target.value })}>
              <option value="">Pilih…</option>
              {(mc?.vendor || []).map((v) => <option key={v?.code} value={v?.name}>{v?.name}</option>)}
            </select>
          </Field>
          <Field label="Jenis Barang">
            <select className="gm-select" value={form.jenis} onChange={(e) => setForm({ ...form, jenis: e.target.value })}>
              <option value="">Pilih…</option>
              {(mc?.jenis || []).map((v) => <option key={v?.code} value={v?.name}>{v?.name}</option>)}
            </select>
          </Field>
          <Field label="Brand">
            <select className="gm-select" value={form.brand} onChange={(e) => setForm({ ...form, brand: e.target.value })}>
              <option value="">Pilih…</option>
              {(mc?.brand || []).map((v) => <option key={v?.code} value={v?.name}>{v?.name}</option>)}
            </select>
          </Field>
          <Field label="Bahan">
            <select className="gm-select" value={form.bahan} onChange={(e) => setForm({ ...form, bahan: e.target.value })}>
              <option value="">Pilih…</option>
              {(mc?.bahan || []).map((v) => <option key={v?.code} value={v?.name}>{v?.name}</option>)}
            </select>
          </Field>
          <Field label="Warna & Artikel">
            <select className="gm-select" value={form.warna} onChange={(e) => setForm({ ...form, warna: e.target.value })}>
              <option value="">Pilih…</option>
              {(mc?.warna || []).map((v) => <option key={v?.code} value={v?.name}>{v?.name}</option>)}
            </select>
          </Field>
          <Field label="Ukuran">
            <select className="gm-select" value={form.ukuran} onChange={(e) => setForm({ ...form, ukuran: e.target.value })}>
              <option value="">Pilih…</option>
              {(mc?.ukuran || []).map((v) => <option key={v?.code} value={v?.name}>{v?.name}</option>)}
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
              {(filtered || []).length === 0 && <EmptyRow colSpan={11} text="Belum ada produk." />}
              {(filtered || []).map((p) => {
                const s = (stockMap?.[p?.kode] || {}) || { masuk: 0, keluar: 0, returOk: 0, reject: 0 };
                const akhir = stokAkhir(p?.kode);
                return (
                  <tr key={p?.kode}>
                    <td style={{ fontFamily: "monospace", fontWeight: 700, color: C.primary }}>{p?.kode}</td>
                    <td style={{ whiteSpace: "normal", minWidth: 220 }}>{p?.nama}</td>
                    <td><span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}><span style={{ width: 9, height: 9, borderRadius: "50%", background: WARNA_SWATCH[p?.warna] || C.muted, display: "inline-block" }} />{p?.warna}</span></td>
                    <td>{p?.ukuran}</td>
                    <td>{p?.vendor}</td>
                    <td>{fmt(s?.masuk || 0)}</td>
                    <td>{fmt(s?.keluar || 0)}</td>
                    <td>{fmt(s?.returOk || 0)}</td>
                    <td>{fmt(s?.reject || 0)}</td>
                    <td style={{ fontWeight: 700, color: akhir <= 10 ? C.danger : C.success }}>{fmt(akhir)}</td>
                    <td><DeleteBtn onClick={() => removeProduk(p?.kode)} /></td>
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
  const [form, setForm] = useState({ tanggal: todayStr(), kode: "", qty: "", catatan: "" });

  const submit = () => {
    if (!form.kode || !Number(form.qty)) { flash("Pilih produk dan isi qty"); return; }
    const prod = (safeData?.products || []).find((p) => p?.kode === form.kode);
    const entry = { id: uid(), tanggal: form.tanggal, kode: form.kode, qty: Number(form.qty), vendor: prod?.vendor || "", catatan: form.catatan };
    save({ ...safeData, masuk: [entry, ...(safeData?.masuk || [])] });
    setForm({ tanggal: form.tanggal, kode: "", qty: "", catatan: "" });
    flash("Barang masuk dicatat");
  };

  const remove = (id) => save({ ...safeData, masuk: (safeData?.masuk || []).filter((m) => m?.id !== id) });

  return (
    <div>
      <SectionCard title="Catat Barang Masuk">
        <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr 1fr 1.5fr auto", gap: 10, alignItems: "end" }}>
          <Field label="Tanggal"><input className="gm-input" type="date" value={form.tanggal} onChange={(e) => setForm({ ...form, tanggal: e.target.value })} /></Field>
          <Field label="Produk"><ProductSelect products={safeData?.products || []} value={form.kode} onChange={(v) => setForm({ ...form, kode: v })} /></Field>
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
              {(safeData?.masuk || []).length === 0 && <EmptyRow colSpan={7} text="Belum ada catatan barang masuk." />}
              {(safeData?.masuk || []).map((m) => {
                const prod = (safeData?.products || []).find((p) => p?.kode === m?.kode);
                return (
                  <tr key={m?.id}>
                    <td>{m?.tanggal}</td>
                    <td style={{ fontFamily: "monospace", fontWeight: 700, color: C.primary }}>{m?.kode}</td>
                    <td style={{ whiteSpace: "normal", minWidth: 200 }}>{prod?.nama || "—"}</td>
                    <td>{fmt(m?.qty || 0)}</td>
                    <td>{m?.vendor}</td>
                    <td>{m?.catatan}</td>
                    <td><DeleteBtn onClick={() => remove(m?.id)} /></td>
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
  const safeData = normalizeData(data || {});
  const emptyForm = { tanggal: todayStr(), kodePesanan: "", shift: "", host: "", kode: "", qty: "", toko: "Shopee", cod: "COD", affiliate: false, namaAffiliate: "", pengiriman: "", ongkir: "", catatan: "" };
  const [form, setForm] = useState(emptyForm);
  const [monthFilter, setMonthFilter] = useState("SEMUA");

  const karyawanList = (safeData?.karyawanMaster || []).map((k) => k?.nama).filter(Boolean) || [];

  const submit = () => {
    if (!form.kode || !Number(form.qty)) { flash("Pilih produk dan isi qty"); return; }
    const sisa = stokAkhir(form.kode);
    if (Number(form.qty) > sisa) { flash(`Stok tidak cukup (tersisa ${sisa})`); return; }
    if (form.affiliate && (!form.pengiriman || !form.ongkir)) { flash("Lengkapi pengiriman & ongkir untuk pesanan affiliate"); return; }
    const entry = { id: uid(), ...form, qty: Number(form.qty), cair: false };
    save({ ...safeData, keluar: [entry, ...(safeData?.keluar || [])] });
    setForm({ ...emptyForm, tanggal: form.tanggal });
    flash("Barang keluar dicatat");
    // Auto-add to bukuKas ketika barang keluar dicatat (Debit otomatis)
    // Tapi Debit akan ditambah ketika status berubah menjadi "Sudah Cair"
  };

  const remove = (id) => save({ ...safeData, keluar: (safeData?.keluar || []).filter((k) => k?.id !== id) });

  const months = ["SEMUA", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"];
  const rows = (safeData?.keluar || []).filter((k) => monthFilter === "SEMUA" || monthOf(k?.tanggal) === monthFilter);
  const totalQty = rows.reduce((s, k) => s + (Number(k?.qty) || 0), 0);

  return (
    <div>
      <SectionCard title="Catat Barang Keluar" subtitle="Stok akan otomatis berkurang">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px,1fr))", gap: 10 }}>
          <Field label="Tanggal"><input className="gm-input" type="date" value={form.tanggal} onChange={(e) => setForm({ ...form, tanggal: e.target.value })} /></Field>
          <Field label="Kode Pesanan"><input className="gm-input" value={form.kodePesanan} onChange={(e) => setForm({ ...form, kodePesanan: e.target.value })} /></Field>
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
          <Field label="Produk"><ProductSelect products={safeData?.products || []} value={form.kode} onChange={(v) => setForm({ ...form, kode: v })} /></Field>
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
            <thead><tr><th>Tanggal</th><th>Kode Pesanan</th><th>Kode</th><th>Produk</th><th>Qty</th><th>Toko</th><th>COD</th><th>Admin/Host</th><th>Catatan</th><th>Cair</th><th></th></tr></thead>
            <tbody>
              {rows?.length === 0 && <EmptyRow colSpan={11} text="Belum ada catatan barang keluar." />}
              {(rows || []).map((k) => {
                const prod = (safeData?.products || []).find((p) => p?.kode === k?.kode);
                return (
                  <tr key={k?.id}>
                    <td>{k?.tanggal}</td>
                    <td>{k?.kodePesanan}</td>
                    <td style={{ fontFamily: "monospace", fontWeight: 700, color: C.primary }}>{k?.kode}</td>
                    <td style={{ whiteSpace: "normal", minWidth: 180 }}>{prod?.nama || "—"}</td>
                    <td>{fmt(k?.qty)}</td>
                    <td>{k?.toko}</td>
                    <td>{k?.cod}</td>
                    <td style={{ fontSize: 11.5, whiteSpace: "normal", minWidth: 140 }}>
                      {k?.host && <div>{k?.host} (Host)</div>}
                      {k?.shift && <div>{k?.shift} (Admin)</div>}
                      {!k?.host && !k?.shift && "—"}
                    </td>
                    <td style={{ whiteSpace: "normal", minWidth: 100 }}>{k?.catatan}</td>
                    <td>
                      <button
                        onClick={() => {
                          const nextKeluar = (safeData?.keluar || []).map((x) => x?.id === k?.id ? { ...x, cair: !x?.cair } : x);
                          const nextBukuKas = [...(safeData?.bukuKas || [])];
                          // Jika berubah menjadi cair, tambahkan ke bukuKas sebagai Debit
                          if (!k?.cair) {
                            nextBukuKas.unshift({
                              id: uid(),
                              tanggal: todayStr(),
                              jenis: "Debit",
                              nominal: 0, // Nominal fleksibel, user input
                              keterangan: `Pembayaran pesanan ${k?.kodePesanan || k?.id?.slice(0, 6)}`,
                            });
                          }
                          save({ ...safeData, keluar: nextKeluar, bukuKas: nextBukuKas });
                        }}
                        style={{ fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 999, border: "none", cursor: "pointer", background: k?.cair ? C.success : C.surfaceAlt, color: k?.cair ? "#fff" : C.muted }}
                      >
                        {k?.cair ? "Sudah" : "Belum"}
                      </button>
                    </td>
                    <td>
                      <button
                        onClick={() => printViaBluetooth(k, "keluar")}
                        title="Cetak ke Bluetooth Printer"
                        style={{ background: "transparent", border: "none", cursor: "pointer", color: C.primary, padding: 4 }}
                      >
                        <Printer size={14} />
                      </button>
                    </td>
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
    save({ ...safeData, retur: [{ id: uid(), ...form, qty: Number(form.qty) }, ...(safeData?.retur || [])] });
    setForm({ tanggal: form.tanggal, kodePesanan: "", kode: "", kodeRetur: "", qty: "", sesuai: "Sesuai", ekspedisi: "", catatan: "" });
    flash("Retur dicatat");
  };
  const remove = (id) => save({ ...safeData, retur: (safeData?.retur || []).filter((r) => r?.id !== id) });

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
          <Field label="Produk"><ProductSelect products={safeData?.products || []} value={form.kode} onChange={(v) => setForm({ ...form, kode: v })} /></Field>
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
              {(safeData?.retur || []).length === 0 && <EmptyRow colSpan={8} text="Belum ada catatan retur." />}
              {(safeData?.retur || []).map((r) => (
                <tr key={r?.id}>
                  <td>{r?.tanggal}</td>
                  <td>{r?.kodePesanan}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: C.primary }}>{r?.kode}</td>
                  <td>{fmt(r?.qty || 0)}</td>
                  <td><span style={{ fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 999, background: r?.sesuai === "Sesuai" ? "#E4EFE4" : "#F5E1DC", color: r?.sesuai === "Sesuai" ? C.success : C.danger }}>{r?.sesuai}</span></td>
                  <td>{r?.ekspedisi}</td>
                  <td style={{ whiteSpace: "normal", minWidth: 140 }}>{r?.catatan}</td>
                  <td><DeleteBtn onClick={() => remove(r?.id)} /></td>
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
  const emptyForm = { tanggal: todayStr(), kode: "", qty: "", toko: "Shopee", catatan: "" };
  const [form, setForm] = useState(emptyForm);

  const submit = () => {
    if (!form.kode || !Number(form.qty)) { flash("Pilih produk dan isi qty"); return; }
    save({ ...safeData, reject: [{ id: uid(), ...form, qty: Number(form.qty) }, ...(safeData?.reject || [])] });
    setForm({ ...emptyForm, tanggal: form.tanggal });
    flash("Reject dicatat");
  };
  const remove = (id) => save({ ...safeData, reject: (safeData?.reject || []).filter((r) => r?.id !== id) });

  return (
    <div>
      <SectionCard title="Catat Reject" subtitle="Barang rusak / cacat akan mengurangi stok akhir">
        <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr 1fr 1fr auto", gap: 10, alignItems: "end" }}>
          <Field label="Tanggal"><input className="gm-input" type="date" value={form.tanggal} onChange={(e) => setForm({ ...form, tanggal: e.target.value })} /></Field>
          <Field label="Produk"><ProductSelect products={safeData?.products || []} value={form.kode} onChange={(v) => setForm({ ...form, kode: v })} /></Field>
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
              {(safeData?.reject || []).length === 0 && <EmptyRow colSpan={7} text="Belum ada catatan reject." />}
              {(safeData?.reject || []).map((r) => {
                const prod = (safeData?.products || []).find((p) => p?.kode === r?.kode);
                return (
                  <tr key={r?.id}>
                    <td>{r?.tanggal}</td>
                    <td style={{ fontFamily: "monospace", fontWeight: 700, color: C.primary }}>{r?.kode}</td>
                    <td style={{ whiteSpace: "normal", minWidth: 180 }}>{prod?.nama || "—"}</td>
                    <td>{fmt(r?.qty || 0)}</td>
                    <td>{r?.toko}</td>
                    <td style={{ whiteSpace: "normal" }}>{r?.catatan}</td>
                    <td><DeleteBtn onClick={() => remove(r?.id)} /></td>
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
  const safeData = normalizeData(data || {});
  const totalKeluar = (safeData?.keluar || []).reduce((s, k) => s + (Number(k?.qty) || 0), 0);
  const totalCair = (safeData?.keluar || []).filter((k) => k?.cair).reduce((s, k) => s + (Number(k?.qty) || 0), 0);
  const belum = (safeData?.keluar || []).filter((k) => !k?.cair);

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
              {(belum || []).length === 0 && <EmptyRow colSpan={6} text="Semua pesanan sudah dicairkan." />}
              {(belum || []).map((k) => (
                <tr key={k?.id}>
                  <td>{k?.tanggal}</td>
                  <td>{k?.kodePesanan}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: C.primary }}>{k?.kode}</td>
                  <td>{fmt(k?.qty || 0)}</td>
                  <td>{k?.toko}</td>
                  <td>
                    <button className="gm-btn" style={{ padding: "5px 11px", fontSize: 12 }} onClick={() => { save({ ...safeData, keluar: (safeData?.keluar || []).map((x) => x?.id === k?.id ? { ...x, cair: true } : x) }); flash("Ditandai sudah cair"); }}>
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

// ---------- KEUANGAN & GAJI TIM (NEW TAB) ----------
function KeuanganTab({ data, save, flash }) {
  const safeData = normalizeData(data || {});
  const [panel, setPanel] = useState("gaji-dinamis");
  const [formKasbon, setFormKasbon] = useState({
    tanggal: todayStr(),
    nama: "",
    nominal: "",
    keterangan: "",
  });

  // State untuk Pengaturan Gaji Dinamis per Karyawan
  const [gajiDinamis, setGajiDinamis] = useState({});

  useEffect(() => {
    // Inisialisasi form gaji dinamis dari data yang tersimpan
    const init = {};
    (safeData?.karyawanMaster || []).forEach((emp) => {
      init[emp?.nama] = {
        gajiPokok: 0,
        komisiPerPcs: 0,
      };
    });
    setGajiDinamis(init);
  }, [safeData?.karyawanMaster]);

  const karyawanList = (safeData?.karyawanMaster || []).map((k) => k?.nama).filter(Boolean) || [];

  const submitKasbon = () => {
    if (!formKasbon.nama.trim() || !Number(formKasbon.nominal)) {
      flash("Isi nama karyawan dan nominal kasbon");
      return;
    }
    const entry = {
      id: uid(),
      tanggal: formKasbon.tanggal,
      nama: formKasbon.nama,
      nominal: Number(formKasbon.nominal),
      keterangan: formKasbon.keterangan,
    };
    save({
      ...safeData,
      kasbon: [entry, ...(safeData?.kasbon || [])],
      bukuKas: [
        {
          id: uid(),
          tanggal: formKasbon.tanggal,
          jenis: "Kredit",
          nominal: Number(formKasbon.nominal),
          keterangan: `Kasbon karyawan ${formKasbon.nama}`,
        },
        ...(safeData?.bukuKas || []),
      ],
    });
    setFormKasbon({ tanggal: todayStr(), nama: "", nominal: "", keterangan: "" });
    flash("Kasbon dicatat dan otomatis masuk ke Buku Kas");
  };

  const removeKasbon = (id) => {
    save({ ...safeData, kasbon: (safeData?.kasbon || []).filter((k) => k?.id !== id) });
  };

  // Simpan pengaturan gaji dinamis
  const saveSalarySettings = () => {
    // Validasi ada minimal satu karyawan dengan pengaturan
    const hasSettings = Object.values(gajiDinamis || {}).some((g) => g?.gajiPokok > 0 || g?.komisiPerPcs > 0);
    if (!hasSettings) {
      flash("Isi minimal satu karyawan dengan nilai gaji");
      return;
    }
    save({
      ...safeData,
      gajiDinamis: gajiDinamis,
    });
    flash("Pengaturan gaji dinamis berhasil disimpan");
  };

  // Hitung slip gaji
  const salarySlips = calculateSalarySlips(safeData);

  const printSlip = (slip) => {
    const html = `
      <html><head><meta charset="utf-8"><title>Slip Gaji - ${slip?.nama}</title></head>
      <body style="font-family: Arial; padding: 20px; max-width: 600px;">
        <h2>SLIP GAJI - NIKI DZIYAB</h2>
        <p><strong>Nama:</strong> ${slip?.nama || "—"}</p>
        <p><strong>Posisi:</strong> ${slip?.role || "—"}</p>
        <p><strong>Tanggal:</strong> ${todayStr()}</p>
        <hr/>
        <p><strong>Qty Barang Keluar:</strong> ${fmt(slip?.qtyKeluar || 0)}</p>
        <p><strong>Gaji Bruto:</strong> Rp ${fmt(slip?.bruto || 0)}</p>
        <p><strong>Potongan Kasbon:</strong> Rp ${fmt(slip?.kasbon || 0)}</p>
        <hr/>
        <p style="font-size: 18px;"><strong>Sisa Gaji Bersih: Rp ${fmt(slip?.bersih || 0)}</strong></p>
        <hr/>
        <p style="font-size: 10px; color: #666;">Dicetak: ${new Date().toLocaleString("id-ID")}</p>
      </body></html>
    `;
    downloadBlob(html, `Slip-Gaji-${slip?.nama}-${todayStr()}.html`, "text/html");
    flash(`Slip gaji ${slip?.nama} diunduh`);
  };

  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 20, flexWrap: "wrap" }}>
        <button
          onClick={() => setPanel("gaji-dinamis")}
          style={{
            padding: "10px 16px",
            borderRadius: 8,
            fontSize: 13,
            fontWeight: 600,
            cursor: "pointer",
            border: `2px solid ${panel === "gaji-dinamis" ? C.primary : C.border}`,
            background: panel === "gaji-dinamis" ? C.primary : "#fff",
            color: panel === "gaji-dinamis" ? "#fff" : C.text,
          }}
        >
          ⚙️ Pengaturan Gaji Manual
        </button>
        <button
          onClick={() => setPanel("kasbon")}
          style={{
            padding: "10px 16px",
            borderRadius: 8,
            fontSize: 13,
            fontWeight: 600,
            cursor: "pointer",
            border: `2px solid ${panel === "kasbon" ? C.primary : C.border}`,
            background: panel === "kasbon" ? C.primary : "#fff",
            color: panel === "kasbon" ? "#fff" : C.text,
          }}
        >
          💰 Form & Histori Kasbon
        </button>
        <button
          onClick={() => setPanel("slip-gaji")}
          style={{
            padding: "10px 16px",
            borderRadius: 8,
            fontSize: 13,
            fontWeight: 600,
            cursor: "pointer",
            border: `2px solid ${panel === "slip-gaji" ? C.primary : C.border}`,
            background: panel === "slip-gaji" ? C.primary : "#fff",
            color: panel === "slip-gaji" ? "#fff" : C.text,
          }}
        >
          📋 Rekap Slip Gaji Tim
        </button>
      </div>

      {/* PANEL 1: Pengaturan Gaji Dinamis */}
      {panel === "gaji-dinamis" && (
        <SectionCard title="Pengaturan Gaji Manual Per Karyawan" subtitle="Atur gaji pokok dan komisi untuk setiap karyawan">
          {karyawanList?.length === 0 ? (
            <div style={{ fontSize: 13, color: C.muted, textAlign: "center", padding: 20 }}>
              Belum ada data karyawan. Silakan tambah karyawan di tab "⚙️ Kode Master" terlebih dahulu.
            </div>
          ) : (
            <>
              <div style={{ overflowX: "auto", marginBottom: 16 }}>
                <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr>
                      <th>Nama Karyawan</th>
                      <th>Gaji Pokok (Rp)</th>
                      <th>Komisi per Pcs (Rp)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {karyawanList.map((nama) => (
                      <tr key={nama}>
                        <td style={{ fontWeight: 600 }}>{nama}</td>
                        <td>
                          <input
                            className="gm-input"
                            type="number"
                            min="0"
                            value={gajiDinamis?.[nama]?.gajiPokok || 0}
                            onChange={(e) => {
                              setGajiDinamis({
                                ...gajiDinamis,
                                [nama]: { ...(gajiDinamis?.[nama] || {}), gajiPokok: Number(e.target.value) },
                              });
                            }}
                            placeholder="0"
                            style={{ width: "100%" }}
                          />
                        </td>
                        <td>
                          <input
                            className="gm-input"
                            type="number"
                            min="0"
                            value={gajiDinamis?.[nama]?.komisiPerPcs || 0}
                            onChange={(e) => {
                              setGajiDinamis({
                                ...gajiDinamis,
                                [nama]: { ...(gajiDinamis?.[nama] || {}), komisiPerPcs: Number(e.target.value) },
                              });
                            }}
                            placeholder="0"
                            style={{ width: "100%" }}
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <button className="gm-btn" onClick={saveSalarySettings}>
                <Save size={15} /> Simpan Pengaturan Gaji
              </button>
            </>
          )}
        </SectionCard>
      )}

      {/* PANEL 2: Form & Histori Kasbon */}
      {panel === "kasbon" && (
        <div>
          <SectionCard title="Tambah Kasbon Karyawan" subtitle="Catat penarikan kasbon dari karyawan">
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, alignItems: "end" }}>
              <Field label="Tanggal">
                <input
                  className="gm-input"
                  type="date"
                  value={formKasbon.tanggal}
                  onChange={(e) => setFormKasbon({ ...formKasbon, tanggal: e.target.value })}
                />
              </Field>
              <Field label="Nama Karyawan">
                <select
                  className="gm-select"
                  value={formKasbon.nama}
                  onChange={(e) => setFormKasbon({ ...formKasbon, nama: e.target.value })}
                >
                  <option value="">Pilih Karyawan…</option>
                  {karyawanList.map((k) => <option key={k} value={k}>{k}</option>)}
                </select>
              </Field>
              <Field label="Nominal (Rp)">
                <input
                  className="gm-input"
                  type="number"
                  min="0"
                  value={formKasbon.nominal}
                  onChange={(e) => setFormKasbon({ ...formKasbon, nominal: e.target.value })}
                  placeholder="0"
                />
              </Field>
              <Field label="Keterangan">
                <input
                  className="gm-input"
                  value={formKasbon.keterangan}
                  onChange={(e) => setFormKasbon({ ...formKasbon, keterangan: e.target.value })}
                  placeholder="Opsional"
                />
              </Field>
              <button className="gm-btn" onClick={submitKasbon}>
                <Plus size={15} /> Simpan Kasbon
              </button>
            </div>
          </SectionCard>

          <SectionCard title="Histori Kasbon">
            <div style={{ overflowX: "auto" }}>
              <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr>
                    <th>Tanggal</th>
                    <th>Nama Karyawan</th>
                    <th>Nominal</th>
                    <th>Keterangan</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {(safeData?.kasbon || []).length === 0 && <EmptyRow colSpan={5} text="Belum ada catatan kasbon." />}
                  {(safeData?.kasbon || []).map((k) => (
                    <tr key={k?.id}>
                      <td>{k?.tanggal}</td>
                      <td style={{ fontWeight: 500 }}>{k?.nama}</td>
                      <td style={{ fontWeight: 700, color: C.secondary }}>Rp {fmt(k?.nominal || 0)}</td>
                      <td>{k?.keterangan}</td>
                      <td>
                        <DeleteBtn onClick={() => removeKasbon(k?.id)} />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </SectionCard>
        </div>
      )}

      {/* PANEL 3: Rekap Slip Gaji Otomatis */}
      {panel === "slip-gaji" && (
        <SectionCard title="Rekap Slip Gaji Tim" subtitle="Slip gaji otomatis berdasarkan transaksi Barang Keluar">
          <div style={{ overflowX: "auto" }}>
            <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  <th>Nama</th>
                  <th>Posisi</th>
                  <th>Qty Jualan</th>
                  <th>Gaji Bruto</th>
                  <th>Potongan Kasbon</th>
                  <th>Sisa Bersih</th>
                  <th>Cetak</th>
                </tr>
              </thead>
              <tbody>
                {(salarySlips || []).length === 0 && (
                  <EmptyRow colSpan={7} text="Belum ada data karyawan dari transaksi Barang Keluar." />
                )}
                {(salarySlips || []).map((slip) => (
                  <tr key={slip?.nama}>
                    <td style={{ fontWeight: 600 }}>{slip?.nama}</td>
                    <td style={{ fontSize: 12, color: C.muted }}>{slip?.role}</td>
                    <td style={{ fontWeight: 500 }}>{fmt(slip?.qtyKeluar || 0)}</td>
                    <td style={{ fontWeight: 600, color: C.primary }}>Rp {fmt(slip?.bruto || 0)}</td>
                    <td style={{ fontWeight: 600, color: C.danger }}>Rp {fmt(slip?.kasbon || 0)}</td>
                    <td style={{ fontWeight: 700, color: C.success }}>Rp {fmt(slip?.bersih || 0)}</td>
                    <td>
                      <button
                        onClick={() => printSlip(slip)}
                        title="Cetak Slip Gaji"
                        style={{
                          background: "transparent",
                          border: "none",
                          cursor: "pointer",
                          color: C.accent,
                          padding: 4,
                          fontSize: 12,
                          fontWeight: 600,
                        }}
                      >
                        🖨️ Cetak
                      </button>
                    </td>
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

// ---------- Helper Functions untuk Keuangan ----------
function calculateSalarySlips(data) {
  const safeData = normalizeData(data || {});
  const slips = [];
  const employeeMap = {};

  // Kumpulkan semua karyawan dari karyawanMaster
  (safeData?.karyawanMaster || []).forEach((emp) => {
    const namaEmp = String(emp?.nama || "").trim();
    if (!namaEmp) return;
    
    employeeMap[namaEmp] = {
      nama: namaEmp,
      role: emp?.peran || "Tim",
      qtyKeluar: 0,
      gajiPokok: 0,
      komisiPerPcs: 0,
    };
  });

  // Hitung qty keluar per karyawan dari transaksi
  (safeData?.keluar || []).forEach((k) => {
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

  // Hitung kasbon per karyawan
  const kasbonMap = {};
  (safeData?.kasbon || []).forEach((k) => {
    const nama = String(k?.nama || "").trim();
    if (!kasbonMap[nama]) kasbonMap[nama] = 0;
    kasbonMap[nama] += Number(k?.nominal) || 0;
  });

  // Ambil pengaturan gaji dinamis jika ada
  const gajiDinamis = data?.gajiDinamis || {};

  // Buat slip gaji
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

// ---------- Fungsi Printer Bluetooth Thermal ----------
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
      Kode: ${item?.kode || "—"}
      Qty: ${item?.qty || 0}
      Pesanan: ${item?.kodePesanan || "—"}
      Toko: ${item?.toko || "—"}
      ===================================
      Cetak otomatis sistem NIKI DZIYAB
      ===================================
    `;

    // Konversi ke bytes dan kirim
    const encoder = new TextEncoder();
    const data = encoder.encode(receiptText);
    await characteristic.writeValue(data);

    alert("Struk berhasil dicetak ke printer Bluetooth!");
  } catch (error) {
    console.error("Error Bluetooth Printer:", error);
    alert("Gagal terhubung ke printer Bluetooth. Pastikan printer dalam jangkauan dan mode pairing aktif.");
  }
}

// ---------- MASTER BUKU KAS (FIXED 📊) ----------
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
      flash("Masukkan nominal");
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
      bukuKas: [entry, ...(safeData?.bukuKas || [])],
    });
    setFormJurnal({ tanggal: todayStr(), jenis: "Debit", nominal: "", keterangan: "" });
    flash("Jurnal kas berhasil dicatat");
  };

  const removeJurnal = (id) => {
    save({ ...safeData, bukuKas: (safeData?.bukuKas || []).filter((b) => b?.id !== id) });
  };

  // 1. Hitung total Debit dan Kredit dari keseluruhan data
  const totalDebit = (safeData?.bukuKas || [])
    .filter((b) => b?.jenis === "Debit")
    .reduce((s, b) => s + (Number(b?.nominal) || 0), 0);

  const totalKredit = (safeData?.bukuKas || [])
    .filter((b) => b?.jenis === "Kredit")
    .reduce((s, b) => s + (Number(b?.nominal) || 0), 0);

  const saldoAkhir = totalDebit - totalKredit;

  // 2. Urutkan data dari tanggal TERTUA ke TERBARU untuk menghitung saldo berjalan secara kronologis
  const sortedKronologis = [...(safeData?.bukuKas || [])].sort(
    (a, b) => new Date(a?.tanggal) - new Date(b?.tanggal)
  );

  let akumulasiSaldo = 0;
  const kronologisWithSaldo = sortedKronologis.map((b) => {
    const nominalItem = Number(b?.nominal) || 0;
    if (b?.jenis === "Debit") {
      akumulasiSaldo += nominalItem;
    } else {
      akumulasiSaldo -= nominalItem;
    }
    return {
      ...b,
      saldoKumulatif: akumulasiSaldo,
    };
  });

  // 3. Balik kembali urutannya agar di tabel memunculkan transaksi TERBARU di posisi paling atas
  const bukuKasWithSaldo = [...kronologisWithSaldo].reverse();

  return (
    <div>
      {/* KOTAK REKAP MUTASI */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 12, marginBottom: 20 }}>
        <div className="gm-card" style={{ padding: 18 }}>
          <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600, textTransform: "uppercase", letterSpacing: ".03em" }}>Total Debit (Uang Masuk)</div>
          <div className="gm-h1" style={{ fontSize: 26, fontWeight: 700, color: C.success, marginTop: 8 }}>Rp {fmt(totalDebit)}</div>
        </div>
        <div className="gm-card" style={{ padding: 18 }}>
          <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600, textTransform: "uppercase", letterSpacing: ".03em" }}>Total Kredit (Uang Keluar)</div>
          <div className="gm-h1" style={{ fontSize: 26, fontWeight: 700, color: C.danger, marginTop: 8 }}>Rp {fmt(totalKredit)}</div>
        </div>
        <div className="gm-card" style={{ padding: 18 }}>
          <div style={{ fontSize: 11.5, color: C.muted, fontWeight: 600, textTransform: "uppercase", letterSpacing: ".03em" }}>Saldo Akhir Kas Toko</div>
          <div className="gm-h1" style={{ fontSize: 26, fontWeight: 700, color: saldoAkhir >= 0 ? C.accent : C.danger, marginTop: 8 }}>Rp {fmt(saldoAkhir)}</div>
        </div>
      </div>

      {/* FORM INPUT MANUAL JURNAL KAS */}
      <SectionCard title="Form Input Manual Jurnal Kas" subtitle="Catat biaya operasional atau transaksi kas tak terduga">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: 12, alignItems: "end" }}>
          <Field label="Tanggal">
            <input
              className="gm-input"
              type="date"
              value={formJurnal.tanggal}
              onChange={(e) => setFormJurnal({ ...formJurnal, tanggal: e.target.value })}
            />
          </Field>
          <Field label="Jenis Transaksi">
            <select
              className="gm-select"
              value={formJurnal.jenis}
              onChange={(e) => setFormJurnal({ ...formJurnal, jenis: e.target.value })}
            >
              <option value="Debit">Debit (Uang Masuk)</option>
              <option value="Kredit">Kredit (Uang Keluar)</option>
            </select>
          </Field>
          <Field label="Nominal (Rp)">
            <input
              className="gm-input"
              type="number"
              min="0"
              value={formJurnal.nominal}
              onChange={(e) => setFormJurnal({ ...formJurnal, nominal: e.target.value })}
              placeholder="0"
            />
          </Field>
          <Field label="Keterangan">
            <input
              className="gm-input"
              value={formJurnal.keterangan}
              onChange={(e) => setFormJurnal({ ...formJurnal, keterangan: e.target.value })}
              placeholder="cth: Biaya air / Hutang supplier"
            />
          </Field>
          <button className="gm-btn" onClick={submitJurnal}>
            <Plus size={15} /> Catat Jurnal
          </button>
        </div>
      </SectionCard>

      {/* TABEL REKENING KORAN MUTASI UTUH */}
      <SectionCard title="Rekening Koran & Saldo Berjalan" subtitle="Daftar lengkap semua transaksi kas dengan saldo kumulatif">
        <div style={{ overflowX: "auto" }}>
          <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th>No</th>
                <th>Tanggal</th>
                <th>Keterangan Transaksi</th>
                <th>Debit (+)</th>
                <th>Kredit (-)</th>
                <th>Saldo Kumulatif</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {bukuKasWithSaldo.length === 0 && (
                <EmptyRow colSpan={7} text="Belum ada transaksi kas." />
              )}
              {bukuKasWithSaldo.map((b, idx) => (
                <tr key={b?.id}>
                  <td style={{ textAlign: "center", color: C.muted, fontSize: 12 }}>{idx + 1}</td>
                  <td>{b?.tanggal}</td>
                  <td style={{ whiteSpace: "normal", minWidth: 200, fontSize: 12.5 }}>{b?.keterangan}</td>
                  <td style={{ fontWeight: 700, color: b?.jenis === "Debit" ? C.success : "inherit" }}>
                    {b?.jenis === "Debit" ? `Rp ${fmt(b?.nominal || 0)}` : "—"}
                  </td>
                  <td style={{ fontWeight: 700, color: b?.jenis === "Kredit" ? C.danger : "inherit" }}>
                    {b?.jenis === "Kredit" ? `Rp ${fmt(b?.nominal || 0)}` : "—"}
                  </td>
                  <td style={{ fontWeight: 700, color: C.primary, fontSize: 12.5 }}>Rp {fmt(b?.saldoKumulatif || 0)}</td>
                  <td>
                    <DeleteBtn onClick={() => removeJurnal(b?.id)} />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </SectionCard>

      {/* INFO SINKRONISASI OTOMATIS */}
      <SectionCard title="ℹ️ Informasi Sinkronisasi Otomatis" subtitle="Fitur-fitur yang terintegrasi dengan Buku Kas">
        <div style={{ fontSize: 13, color: C.text, lineHeight: 1.7 }}>
          <p>
            <strong>✓ Kasbon → Kredit Otomatis:</strong> Setiap kali Anda menambah Kasbon di tab "💰 Keuangan & Gaji Tim", sistem otomatis mencatat transaksi sebagai <strong>Kredit (Uang Keluar)</strong> ke Buku Kas.
          </p>
          <p>
            <strong>✓ Barang Keluar → Debit (Sudah Cair):</strong> Ketika status pesanan di tab "Barang Keluar" diubah menjadi "Sudah Cair", sistem otomatis mencatat sebagai <strong>Debit (Uang Masuk)</strong> ke Buku Kas.
          </p>
          <p>
            <strong>✓ Perhitungan Otomatis:</strong> Saldo akhir kas toko dihitung langsung dari rumus: <code>Total Debit - Total Kredit</code>.
          </p>
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
  { key: "karyawanMaster", label: "👥 Karyawan & Tim" },
];

function MasterTab({ data, save, flash }) {
  const safeData = normalizeData(data || {});
  const [cat, setCat] = useState("vendor");
  const [form, setForm] = useState({ code: "", name: "", peran: "" });

  // Untuk Karyawan Master
  const [formKaryawan, setFormKaryawan] = useState({ nama: "", peran: "Host Live" });

  const karyawanList = (safeData?.karyawanMaster || []) || [];
  const roleOptions = ["Owner", "Host Live", "Admin Live"];

  // Data untuk kategori biasa (vendor, jenis, brand, dll)
  const listData = cat === "karyawanMaster" ? [] : ((safeData?.masterCode?.[cat] || []) || []);

  const addKaryawan = () => {
    if (!formKaryawan.nama.trim()) {
      flash("Masukkan nama karyawan");
      return;
    }
    if (karyawanList.some((k) => k?.nama?.toLowerCase() === formKaryawan.nama.trim().toLowerCase())) {
      flash("Nama karyawan sudah terdaftar");
      return;
    }
    const newEmployee = {
      id: uid(),
      nama: formKaryawan.nama.trim(),
      peran: formKaryawan.peran,
    };
    save({
      ...safeData,
      karyawanMaster: [...karyawanList, newEmployee],
    });
    setFormKaryawan({ nama: "", peran: "Host Live" });
    flash(`Karyawan ${formKaryawan.nama} berhasil ditambahkan`);
  };

  const removeKaryawan = (id) => {
    save({
      ...safeData,
      karyawanMaster: karyawanList.filter((k) => k?.id !== id),
    });
    flash("Karyawan berhasil dihapus");
  };

  const addMaster = () => {
    if (!form.code.trim() || !form.name.trim()) {
      flash("Isi kode dan nama");
      return;
    }
    if (listData.some((x) => x?.code?.toUpperCase() === form.code.trim().toUpperCase())) {
      flash("Kode sudah dipakai di kategori ini");
      return;
    }
    const next = {
      ...safeData,
      masterCode: {
        ...safeData?.masterCode,
        [cat]: [
          ...(listData || []),
          { code: form.code.trim().toUpperCase(), name: form.name.trim().toUpperCase() },
        ],
      },
    };
    save(next);
    setForm({ code: "", name: "", peran: "" });
    flash("Kode master berhasil ditambahkan");
  };

  const removeMaster = (code) => {
    save({
      ...safeData,
      masterCode: {
        ...safeData?.masterCode,
        [cat]: (listData || []).filter((x) => x?.code !== code),
      },
    });
  };

  return (
    <div>
      <SectionCard title="Kode Master & Data Referensi" subtitle="Kombinasi kode-kode ini membentuk sistem pembukuan terintegrasi">
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 16, overflowX: "auto" }}>
          {KATEGORI.map((k) => (
            <button
              key={k.key}
              onClick={() => {
                setCat(k.key);
                setForm({ code: "", name: "", peran: "" });
                setFormKaryawan({ nama: "", peran: "Host Live" });
              }}
              style={{
                padding: "7px 13px",
                borderRadius: 999,
                fontSize: 12.5,
                fontWeight: 600,
                cursor: "pointer",
                border: `1px solid ${cat === k.key ? C.primary : C.border}`,
                background: cat === k.key ? C.primary : "#fff",
                color: cat === k.key ? "#fff" : C.text,
                whiteSpace: "nowrap",
              }}
            >
              {k.label}
            </button>
          ))}
        </div>

        {/* SECTION: KARYAWAN & TIM */}
        {cat === "karyawanMaster" && (
          <div>
            <div style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 13.5, fontWeight: 700, color: C.primaryDark, marginBottom: 12 }}>Tambah Karyawan Baru</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 10, alignItems: "end" }}>
                <Field label="Nama Karyawan">
                  <input
                    className="gm-input"
                    value={formKaryawan.nama}
                    onChange={(e) => setFormKaryawan({ ...formKaryawan, nama: e.target.value })}
                    placeholder="Masukkan nama…"
                  />
                </Field>
                <Field label="Peran / Role">
                  <select
                    className="gm-select"
                    value={formKaryawan.peran}
                    onChange={(e) => setFormKaryawan({ ...formKaryawan, peran: e.target.value })}
                  >
                    {roleOptions.map((role) => (
                      <option key={role} value={role}>
                        {role}
                      </option>
                    ))}
                  </select>
                </Field>
                <button className="gm-btn" onClick={addKaryawan}>
                  <Plus size={15} /> Tambah Karyawan
                </button>
              </div>
            </div>

            <div style={{ overflowX: "auto" }}>
              <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr>
                    <th>Nama Karyawan</th>
                    <th>Peran / Role</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {karyawanList?.length === 0 && (
                    <EmptyRow colSpan={3} text="Belum ada data karyawan. Silakan tambah karyawan baru." />
                  )}
                  {(karyawanList || []).map((k) => (
                    <tr key={k?.id}>
                      <td style={{ fontWeight: 600, fontSize: 13.5 }}>{k?.nama}</td>
                      <td>
                        <span
                          style={{
                            fontSize: 11.5,
                            fontWeight: 600,
                            padding: "4px 9px",
                            borderRadius: 6,
                            background: k?.peran === "Owner" ? "#FEE2E2" : k?.peran === "Host Live" ? "#DBEAFE" : "#F3E8FF",
                            color: k?.peran === "Owner" ? "#7F1D1D" : k?.peran === "Host Live" ? "#0C4A6E" : "#5B21B6",
                          }}
                        >
                          {k?.peran}
                        </span>
                      </td>
                      <td>
                        <DeleteBtn onClick={() => removeKaryawan(k?.id)} />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* SECTION: MASTER KODE BIASA (VENDOR, JENIS, BRAND, BAHAN, WARNA, UKURAN) */}
        {cat !== "karyawanMaster" && (
          <div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr auto", gap: 10, marginBottom: 14 }}>
              <Field label="Kode">
                <input
                  className="gm-input"
                  value={form.code}
                  onChange={(e) => setForm({ ...form, code: e.target.value })}
                  placeholder="cth: B"
                />
              </Field>
              <Field label="Nama">
                <input
                  className="gm-input"
                  value={form.name}
                  onChange={(e) => setForm({ ...form, name: e.target.value })}
                  placeholder="cth: TOKO BARU"
                />
              </Field>
              <div style={{ display: "flex", alignItems: "end" }}>
                <button className="gm-btn" onClick={addMaster}>
                  <Plus size={15} /> Tambah
                </button>
              </div>
            </div>

            <div style={{ overflowX: "auto" }}>
              <table className="gm-table" style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr>
                    <th>Kode</th>
                    <th>Nama</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {(listData || []).length === 0 && <EmptyRow colSpan={3} text="Belum ada data." />}
                  {(listData || []).map((v) => (
                    <tr key={v?.code}>
                      <td style={{ fontFamily: "monospace", fontWeight: 700 }}>{v?.code}</td>
                      <td>{v?.name}</td>
                      <td>
                        <DeleteBtn onClick={() => removeMaster(v?.code)} />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
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
