# SIMPATI — Aset Tetap & Perpindahan Barang

Website statis sederhana untuk **melihat data aset tetap**, **mencatat perpindahan barang**,
**menambah barang**, **menghitung kelayakan biaya perbaikan**, dan **mengunggah data Excel baru**.
Seluruhnya berjalan offline di browser — tanpa internet, CDN, atau proses build.

## Cara membuka

**Cara termudah (disarankan):**

1. Buka folder `D:\aset-tetap`.
2. Klik dua kali **`jalankan-server.bat`**.
3. Browser terbuka otomatis ke **http://127.0.0.1:8765/**.
4. Biarkan jendela hitam server tetap terbuka selama memakai website.
   Tekan `Ctrl+C` di jendela tersebut untuk menghentikan server.

**Cara alternatif:** klik dua kali **`index.html`** atau **`buka-website.bat`**.
Namun saat dibuka lewat `file://`, sebagian browser memblokir penyimpanan data,
sehingga server lokal lebih dianjurkan.

## Fitur

| Halaman | Isi |
|---|---|
| **Ringkasan** | Total nilai aset, jumlah unit, perpindahan bulan ini, rata-rata nilai, grafik nilai & sebaran sub kelompok, perpindahan terbaru, lokasi dengan aset terbanyak. |
| **Data Aset** | 3.200 unit aset. Pencarian (nama/kode/merek/lokasi), filter sub kelompok & lokasi, urutkan, detail aset + riwayatnya, ekspor CSV. |
| **Perpindahan** | Formulir 7 langkah: pilih aset, lokasi asal otomatis, lokasi tujuan, tanggal, petugas, alasan, catatan. Pratinjau, validasi, dan bukti perpindahan yang bisa dicetak. |
| **Perbaikan** | Kalkulator kelayakan perbaikan: `estimasi = harga perolehan × 15%`. Masukkan **biaya perbaikan**, lalu sistem menilai dua syarat. Hasil bisa disimpan, dilihat riwayatnya, dan diekspor CSV. |
| **Penyusutan** | Penyusutan metode **garis lurus**: kolom aset barang, kategori, masa manfaat, harga satuan, penyusutan tahunan & bulanan, umur, akumulasi, dan nilai buku. Bisa dicari, difilter per kategori/status, dan diekspor CSV. |
| **Riwayat** | Semua mutasi terbaru, pencarian & filter alasan, hapus catatan (lokasi dihitung ulang), ekspor CSV. |
| **Tambah Barang** | Formulir aset baru lengkap dengan pratinjau dan validasi. Barang langsung masuk ke dataset aktif dan tersimpan permanen di browser. |
| **Admin** | Panel terkunci PIN untuk mengunggah Excel/CSV, memetakan kolom, mengganti atau menggabungkan data, mengelola PIN, log aktivitas, backup/restore JSON, dan reset dataset. |

### Aturan perbaikan

Aset **memenuhi syarat perbaikan** hanya jika **kedua** kondisi terpenuhi:

1. `biaya perbaikan` **> Rp 4.000.000**, dan
2. `harga perolehan × 15%` **> Rp 4.000.000**.

Nilai yang tepat Rp 4.000.000 dianggap **belum** memenuhi syarat (harus lebih besar).

### Penyusutan garis lurus

Rumus yang dipakai:

```
Penyusutan tahunan = (harga perolehan − nilai residu 0) ÷ masa manfaat
Penyusutan bulanan = penyusutan tahunan ÷ 12
Akumulasi          = penyusutan bulanan × jumlah bulan sejak tanggal perolehan
Nilai buku         = harga perolehan − akumulasi
```

Masa manfaat ditentukan dari **kategori aset** yang mengikuti kelompok sheet pada file sumber:

| Kategori | Masa manfaat |
|---|---|
| ALUM 1 | 4 tahun |
| ALUM 2 | 8 tahun |
| ALKES 1 | 4 tahun |
| ALKES 2 | 8 tahun |
| MOBIL | 8 tahun |
| MOTOR | 4 tahun |
| KSO ALKES 1 | 4 tahun |
| KSO ALKES 2 | 8 tahun |

Umur aset dihitung sejak **tanggal perolehan**. Penyusutan berhenti setelah masa manfaat
habis, sehingga nilai buku tidak pernah negatif. Aset yang kodenya tidak dikenali akan
tampil sebagai **Belum berkategori** (tanpa penyusutan) — pemetaan kategori bawaan ada di
`kategori.js`.

## Panel Admin

- PIN awal: **1234** (dapat diganti di panel admin).
- PIN disimpan sebagai hash di browser, bukan pengamanan tingkat server —
  hanya untuk mencegah perubahan data tanpa sengaja.
- Halaman admin otomatis terkunci lagi saat tab/browser ditutup.

### Unggah data Excel

1. Buka **Admin** dan masukkan PIN.
2. Pilih file **.xlsx** atau **.csv** (bisa juga seret ke kotak unggah).
3. Pilih sheet dan baris header bila perlu — kolom wajib dipetakan otomatis,
   dan bisa diubah manual. Kolom lokasi tanpa judul dikenali dari isi selnya.
4. Periksa pratinjau, jumlah baris valid/dilewati/duplikat.
5. Pilih mode:
   - **Ganti** — dataset aktif diganti seluruhnya.
   - **Gabungkan** — data diperbarui berdasarkan kode aset, aset lain dipertahankan.
6. Klik **Impor data**. Data tersimpan permanen di browser (IndexedDB).

Tersedia juga **Unduh template**, **Backup JSON**, **Restore**, dan
**Kembalikan data bawaan** (3.200 aset).

## Data

- Sumber awal: `LOKTAIL OKTOBER 2026` (sheet **ALL**).
- Jumlah aset bawaan: **3.200 unit** di **242 lokasi**.
- Total harga perolehan: **Rp 45.866.572.278** (± Rp 45,87 miliar).
- Sub kelompok: Alat Medis dan Kesehatan, Alat Umum dan Rumah Tangga,
  Peralatan Komputer, Kendaraan Bermotor, Perangkat Lunak, Lain-lain.

Lima catatan perpindahan awal disediakan sebagai **contoh** (bertanda `demo` di riwayat)
agar dashboard tidak kosong. Hapus lewat **“Reset data perpindahan”** di panel admin.

## Struktur file

```
aset-tetap/
├─ index.html        # kerangka halaman
├─ styles.css        # seluruh tampilan (responsif, tanpa CDN)
├─ app.js            # logika aplikasi (vanilla JavaScript)
├─ data.js           # data aset bawaan hasil ekstraksi Excel
├─ kategori.js       # pemetaan kode aset → kategori (ALUM/ALKES/MOBIL/MOTOR/KSO)
├─ xlsx-reader.js    # pembaca .xlsx/.csv offline (ZIP + inflate + XML)
├─ dataset-store.js  # penyimpanan dataset aktif (IndexedDB)
├─ server.ps1        # server lokal sederhana (PowerShell)
├─ jalankan-server.bat
├─ buka-website.bat
├─ favicon.svg
└─ README.md
```

## Format `data.js`

```js
window.ASSET_DB = {
  meta:   { total, count, source, created },
  lokasi: [["kode", "nama lokasi"], ...],
  sub:    ["sub kelompok", ...],
  kel:    ["kelompok", ...],
  merek:  ["merek", ...],
  aset:   [["ref", "nama", locId, "YYYY-MM-DD", subId, merekId, harga, kelId, "deskripsi"], ...]
};
```

`locId`, `subId`, `merekId`, dan `kelId` adalah indeks pada array masing-masing.

## Catatan penyimpanan

- Dataset aktif (bawaan/unggahan/barang tambahan) disimpan di **IndexedDB**.
- Perpindahan disimpan di **localStorage** (`simpati_state_v1`).
- PIN admin: `simpati_admin_pin_v1`; log aktivitas: `simpati_admin_log_v1`;
  riwayat perbaikan: `simpati_repairs_v1`.
- Data tidak dikirim ke mana pun dan tidak mengubah file Excel asli.
- Menghapus data situs di browser akan menghapus seluruh catatan dan dataset unggahan
  (data bawaan tetap dapat dipulihkan lewat panel admin).
