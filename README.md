# RPS Builder OBE

Aplikasi Streamlit untuk menyusun RPS berbasis OBE dari master kurikulum Excel dan template Word prodi. Aplikasi ini tidak membuat format RPS dari nol, tetapi mengisi placeholder pada template `.docx` agar format resmi prodi tetap dipertahankan.

## Fitur

- Master bawaan dipilih berdasarkan jenjang dan angkatan: D4 2023–2026 serta D3 2024–2026.
- Kode, struktur semester, SKS, dan jam per minggu mengikuti pemetaan tiap angkatan.
- Materi D4 dibentuk dari `Short Silabus Kurikulum D4 PSTE.pdf`, bukan materi generik.
- Upload Excel master kurikulum sendiri tetap tersedia.
- Export Word memakai template default `templates/template_rps_pste_placeholder.docx`.
- Pilih mata kuliah dari `Master_MK`.
- Pratinjau data CPL, IK, short silabus, CPMK, rencana mingguan, dan referensi.
- Editor untuk dosen: dosen pengampu, deskripsi MK, CPMK, referensi, dan tabel rencana mingguan sesuai jumlah pertemuan pada master (D3 17, D4 16).
- Tab `Rencana Mingguan` memakai `st.data_editor` dengan dropdown modalitas, bentuk pembelajaran, metode pembelajaran, dan teknik asesmen.
- Pertemuan 9 digunakan untuk UTS, sedangkan pertemuan terakhir digunakan untuk UAS atau evaluasi/proyek akhir semester.
- Validasi OBE:
  - setiap CPMK harus punya IK;
  - setiap IK harus punya CPL;
  - total bobot penilaian harus 100%;
  - setiap pertemuan sebaiknya memiliki kemampuan akhir yang direncanakan;
  - kemampuan akhir yang terisi sebaiknya memiliki teknik asesmen;
  - setiap CPL yang dibebankan muncul minimal pada satu CPMK.
  - konflik `Mapping_MK_CPL` dengan jalur `Master_CPMK → Master_IK` menghentikan proses;
  - CPL langsung pada CPMK harus sama dengan CPL induk IK;
  - kode MK duplikat wajib memakai `id_penawaran` lengkap pada seluruh sheet terkait.
- Referensi `CPMK4` lama pada master D3 yang tidak valid dipetakan ke CPMK terakhir
  yang tersedia dan diberi penanda `AUTO_REMAP_REVIEW_DOSEN` untuk pemeriksaan dosen.
- Export:
  - RPS Word sesuai template prodi;
  - RPS Excel;
  - laporan validasi CSV.

## Struktur

```text
.
├── app.py
├── curriculum.py
├── data/
│   └── curriculum_catalog.json
├── requirements.txt
├── README.md
├── templates/
│   └── template_rps_pste_placeholder.docx
├── sample_data/
│   ├── master_rps_d3_pste.xlsx
│   └── master_rps_d4_pste.xlsx
└── outputs/
```

File pada `sample_data` menjadi basis CPL/IK/CPMK. `curriculum_catalog.json` menyimpan struktur dan kode resmi per angkatan serta short silabus D4 yang sudah diekstrak dari dokumen sumber. Master angkatan dibentuk ketika dipilih dan dapat langsung diunduh dari sidebar. Kode yang sama untuk dua penawaran tetap dipertahankan dan dipisahkan dengan `id_penawaran`.

## Instalasi

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Menjalankan

```bash
streamlit run app.py
```

## Format Excel

Workbook harus memiliki sheet berikut:

Nama file Excel bebas. Aplikasi tidak memvalidasi nama file; yang dibaca adalah nama sheet dan kolom di dalam workbook.

- `Master_CPL`: `kode_cpl`, `deskripsi_cpl`
- `Master_IK`: `kode_ik`, `deskripsi_ik`, `kode_cpl`
- `Master_MK`: `kode_mk`, `nama_mk`, `nama_prodi`, `semester`, `sks_teori`, `sks_praktek`, `jenis_mk`; `id_penawaran` digunakan bila satu kode MK dipakai oleh lebih dari satu penawaran.
- `Mapping_MK_CPL`: `kode_mk`, `kode_cpl`
- `Master_CPMK`: `kode_mk`, `kode_cpmk`, `deskripsi_cpmk` atau `rumusan_cpmk`, `kode_ik`, `kode_cpl`
- `Short_Silabus`: `kode_mk`, `deskripsi_mk`, `bahan_kajian`
- `RPS_Pertemuan`: `kode_mk`, `minggu`, `sub_cpmk`, `materi`, `modalitas`, `bentuk_pembelajaran`, `metode`, `pengalaman_belajar`, `teknik_asesmen`, `indikator_penilaian`, `bobot`, `referensi`, `kode_cpmk`
- `Referensi`: `kode_mk`, `referensi`
- Opsional untuk tabel evaluasi Word: `Evaluasi_RPS`, `Asesmen_Mingguan`

Nama kolom akan dinormalisasi menjadi huruf kecil dan spasi/tanda hubung menjadi underscore.
Jika sheet `RPS_Pertemuan` berisi 16 pertemuan, aplikasi mempertahankan 16 pertemuan. Workbook lama D3 yang berisi 17 pertemuan tetap didukung.

Contoh otomatis mencakup mata kuliah PLC, Instrumentasi Industri, dan Proyek Akhir.

Kode CPL dinormalisasi otomatis. Format seperti `CPL01`, `CPL02`, dan `CPL10` akan dibaca sebagai `CPL1`, `CPL2`, dan `CPL10`. Kolom `kode_mk` juga dibaca sebagai teks dan di-strip dari spasi tersembunyi. Jika mapping CPL atau IK mengacu ke kode yang tidak ada di master, aplikasi menampilkan warning pada preview dan laporan validasi.

CPL yang dibebankan pada RPS diambil dari unique `kode_cpl` pada `Master_CPMK` untuk mata kuliah terpilih, lalu di-join ke `Master_CPL`. Semua sheet terkait difilter dengan `id_penawaran` bila tersedia sehingga dua penawaran berkode MK sama tidak tercampur.

## Placeholder Template Word Aman

Template Word default berada di `templates/template_rps_pste_placeholder.docx`. Aplikasi tidak menyediakan upload template dari sidebar. Template disarankan hanya memakai placeholder sederhana, tanpa loop Jinja di dalam tabel Word. Placeholder huruf kecil berikut adalah format utama:

```text
{{ nama_prodi }}
{{ kode_mk }}
{{ nama_mk }}
{{ semester }}
{{ sks_teori }}
{{ sks_praktek }}
{{ total_sks }}
{{ dosen_pengampu }}
{{ deskripsi_mk }}
{{ cpl_text }}
{{ ik_text }}
{{ cpmk_text }}
{{ bahan_kajian }}
{{ rencana_mingguan_text }}
{{ asesmen_text }}
{{ referensi_text }}
```

Placeholder lama dengan huruf besar seperti `{{NAMA_MK}}`, `{{CPL_DIBEBANKAN}}`, dan `{{REFERENSI}}` masih dikirim di context untuk kompatibilitas.

Jangan gunakan `{% for %}` dan `{% endfor %}` di dalam tabel Word karena formatting Word dapat memecah tag Jinja dan memicu error `unknown tag 'endfor'`. Data list seperti CPL, IK, CPMK, asesmen, dan referensi sudah dibentuk sebagai teks multiline oleh aplikasi. Rencana mingguan dibuat sebagai tabel Word menggunakan `python-docx`; kolom `sub_cpmk` diekspor sebagai `Kemampuan akhir yang direncanakan`.
# rps-builder-pste
