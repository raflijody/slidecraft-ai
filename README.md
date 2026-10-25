# SlideCraft AI — Prototype v0.1

Website prototipe untuk menyusun presentasi akademik dan mengekspor file PowerPoint yang dapat diedit.

## Yang sudah ada
- Antarmuka responsif untuk komputer, tablet, dan iPad.
- Form topik, audiens, jumlah slide, gaya visual, arahan tambahan.
- Upload materi sumber `.txt`, `.md`, `.docx`, atau `.pptx`.
- Upload template Canva yang sudah diekspor ke `.pptx`.
- Integrasi API OpenAI melalui backend (API key tidak dikirim ke browser).
- Ekspor hasil sebagai `.pptx` yang bisa diedit.
- Mode fallback untuk membuat draf dasar tanpa API key.
- Instruksi model agar tidak mengarang angka dan referensi.

## Keterbatasan penting versi awal
1. **Template PPTX:** backend mengisi placeholder pada slide yang sudah ada. Paling andal jika template memiliki kotak teks berisi `{{TITLE}}`, `{{BODY}}`, `{{NOTES}}`, `{{VISUAL}}`, dan/atau `{{REFERENCES}}`. Tanpa placeholder, aplikasi memakai heuristik sederhana untuk mengganti kotak teks pertama dan kedua; ini dapat merusak desain atau mengubah elemen yang tidak dimaksud.
2. Jumlah slide template sebaiknya sama dengan jumlah slide yang diminta. Versi ini belum mengkloning layout slide baru secara aman.
3. PDF belum didukung sebagai sumber materi.
4. Model AI bisa salah memahami materi. Pemeriksaan manusia tetap diperlukan; instruksi untuk tidak mengarang bukan jaminan mutlak.
5. Grafik, gambar, tabel kompleks, catatan pembicara, dan layout Canva tidak otomatis dibuat secara cerdas pada versi ini.
6. Font/elemen dari Canva dapat berubah saat ekspor ke PPTX, bergantung pada template dan dukungan PowerPoint.

## Cara menjalankan di komputer

Perlu Python 3.10 atau lebih baru.

### 1. Ekstrak ZIP dan buka folder `slidecraft_ai`

### 2. Buat virtual environment (opsional tetapi disarankan)
macOS/Linux:
```bash
python3 -m venv .venv
source .venv/bin/activate
```
Windows:
```powershell
python -m venv .venv
.venv\\Scripts\\activate
```

### 3. Pasang dependensi
```bash
pip install -r requirements.txt
```

### 4. Atur API key
macOS/Linux:
```bash
export OPENAI_API_KEY="API_KEY_ANDA"
```
Windows PowerShell:
```powershell
$env:OPENAI_API_KEY="API_KEY_ANDA"
```
Opsional: set `OPENAI_MODEL` untuk memilih model yang tersedia di akun API Anda. Pemakaian API dapat dikenai biaya sesuai tarif penyedia.

**Jangan pernah menaruh API key di `static/app.js`, HTML, atau repositori publik.** Untuk deployment, gunakan environment variables/secrets pada server.

### 5. Jalankan server
```bash
uvicorn app:app --reload
```
Buka `http://127.0.0.1:8000`.

## Cara menyiapkan template Canva
1. Buka desain di Canva.
2. Unduh sebagai **Microsoft PowerPoint (.pptx)** jika opsi tersebut tersedia.
3. Untuk slide yang ingin diisi otomatis, ganti teks contoh di kotak teks yang tepat dengan placeholder, misalnya `{{TITLE}}` dan `{{BODY}}`.
4. Buat jumlah slide sesuai jumlah yang akan diminta.
5. Unggah PPTX itu ke aplikasi.

Placeholder akan diganti dengan konten. Pertahankan kotak teks dan desain yang ada; hindari menaruh placeholder di banyak kotak yang tidak semestinya.

## Arsitektur
- Frontend: HTML, CSS, JavaScript.
- Backend: FastAPI.
- PPTX: python-pptx.
- AI: OpenAI API di sisi server.
- Data: tidak ada database pada prototipe; file diproses saat permintaan berlangsung.

## Sebelum dipakai publik
Tambahkan autentikasi, rate limiting, batas ukuran file, pembersihan file sementara, logging aman, kebijakan privasi, pembatasan biaya API, dan pemeriksaan file. Jangan menjadikan prototipe ini layanan publik tanpa pengamanan tambahan.
