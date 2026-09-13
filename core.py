"""Mesin pengolah laporan Excel LMS Edlink untuk ESM Activity Analyzer."""

from io import BytesIO
import re

import pandas as pd


BOBOT = {"Materi": 0.25, "Tugas": 0.40, "Kuis": 0.35}


def _angka(nilai):
    hasil = pd.to_numeric(pd.Series([nilai]), errors="coerce").iloc[0]
    return 0.0 if pd.isna(hasil) else float(hasil)


def _terisi(nilai):
    return not pd.isna(nilai) and str(nilai).strip() != ""


def _jenis_dan_judul(header):
    teks = str(header).strip()
    jenis = "Lainnya"
    for kandidat in ("Materi", "Tugas", "Quiz", "Kuis"):
        if teks.lower().startswith(kandidat.lower()):
            jenis = "Kuis" if kandidat in ("Quiz", "Kuis") else kandidat
            break
    cocok = re.search(r"\((.*)\)", teks)
    return jenis, cocok.group(1).strip() if cocok else teks


def _kelompok_konten(data):
    header_utama = list(data.iloc[4])
    subheader = list(data.iloc[5])
    awal = [i for i in range(2, len(header_utama)) if _terisi(header_utama[i])]
    kelompok = []
    for posisi, start in enumerate(awal):
        end = awal[posisi + 1] if posisi + 1 < len(awal) else len(header_utama)
        jenis, judul = _jenis_dan_judul(header_utama[start])
        kolom = {str(subheader[i]).strip(): i for i in range(start, end) if _terisi(subheader[i])}
        kelompok.append({"jenis": jenis, "judul": judul, "kolom": kolom})
    return kelompok


def _nilai_baris(baris, kolom, nama):
    indeks = kolom.get(nama)
    return None if indeks is None or indeks >= len(baris) else baris.iloc[indeks]


def _baca_sheet(data, nama_sheet):
    if data.shape[0] < 7 or data.shape[1] < 2:
        return [], []
    kelompok = _kelompok_konten(data)
    nomor_sesi = data.iloc[2, 1] if data.shape[1] > 1 else None
    label_sesi = f"Sesi {nomor_sesi}" if _terisi(nomor_sesi) else nama_sheet.strip() or "Tanpa Sesi"
    catatan = []
    peserta = data.iloc[6:].copy()
    peserta = peserta[peserta.iloc[:, 0].astype(str).str.contains("@", na=False)]

    for _, baris in peserta.iterrows():
        email = str(baris.iloc[0]).strip().lower()
        nama = str(baris.iloc[1]).strip()
        for urutan, item in enumerate(kelompok, start=1):
            kolom = item["kolom"]
            views = _angka(_nilai_baris(baris, kolom, "Views"))
            komentar = _angka(_nilai_baris(baris, kolom, "Komentar"))
            video = _angka(_nilai_baris(baris, kolom, "Video Dilihat"))
            unduhan = _angka(_nilai_baris(baris, kolom, "Dokumen Didownload"))
            waktu_kirim = _nilai_baris(baris, kolom, "Waktu Kirim Tugas")
            nilai_tugas = _nilai_baris(baris, kolom, "Tugas")
            nilai_kuis = _nilai_baris(baris, kolom, "Quiz")
            if nilai_kuis is None:
                nilai_kuis = _nilai_baris(baris, kolom, "Kuis")
            akses = views > 0 or video > 0 or unduhan > 0
            dikumpulkan = item["jenis"] == "Tugas" and (_terisi(waktu_kirim) or _terisi(nilai_tugas))
            ikut_kuis = item["jenis"] == "Kuis" and _terisi(nilai_kuis)
            catatan.append({
                "Email": email, "Nama": nama, "Sheet": nama_sheet, "Sesi": label_sesi,
                "ID Konten": f"{nama_sheet}::{urutan}",
                "Nomor Item": urutan, "Jenis": item["jenis"], "Judul": item["judul"],
                "Views": views, "Komentar": komentar, "Video Dilihat": video,
                "Dokumen Diunduh": unduhan, "Diakses": akses,
                "Tugas Dikumpulkan": dikumpulkan,
                "Waktu Kirim Tugas": waktu_kirim if _terisi(waktu_kirim) else None,
                "Mengikuti Kuis": ikut_kuis,
                "Nilai Kuis": _angka(nilai_kuis) if ikut_kuis else None,
            })
    return catatan, kelompok


def daftar_konten(sumber):
    """Menghasilkan daftar materi, tugas, dan kuis yang dapat dipilih pengguna."""
    if isinstance(sumber, (bytes, bytearray)):
        sumber = BytesIO(sumber)
    excel = pd.ExcelFile(sumber)
    semua = []
    for sheet in excel.sheet_names:
        data = pd.read_excel(excel, sheet_name=sheet, header=None)
        catatan, _ = _baca_sheet(data, sheet)
        semua.extend(catatan)
    if not semua:
        raise ValueError("Struktur data peserta Edlink tidak ditemukan dalam file.")
    detail = pd.DataFrame(semua)
    return detail[["ID Konten", "Sesi", "Jenis", "Judul"]].drop_duplicates().reset_index(drop=True)


def analisis_edlink(sumber, konten_dipilih=None, batas_aktif=75, batas_cukup=50):
    if isinstance(sumber, (bytes, bytearray)):
        sumber = BytesIO(sumber)
    excel = pd.ExcelFile(sumber)
    semua, ringkasan_sesi = [], []
    for sheet in excel.sheet_names:
        data = pd.read_excel(excel, sheet_name=sheet, header=None)
        catatan, kelompok = _baca_sheet(data, sheet)
        semua.extend(catatan)
        if catatan:
            d = pd.DataFrame(catatan)
            ringkasan_sesi.append({
                "Sesi/Sheet": d["Sesi"].iloc[0], "Nama Sheet": sheet,
                "Jumlah Peserta": d["Email"].nunique(), "Jumlah Konten": len(kelompok),
                "Peserta Mengakses": d.loc[d["Diakses"], "Email"].nunique(),
                "Peserta Mengumpulkan Tugas": d.loc[d["Tugas Dikumpulkan"], "Email"].nunique(),
                "Peserta Mengikuti Kuis": d.loc[d["Mengikuti Kuis"], "Email"].nunique(),
            })
    if not semua:
        raise ValueError("Struktur data peserta Edlink tidak ditemukan dalam file.")

    detail = pd.DataFrame(semua)
    if konten_dipilih is not None:
        detail = detail[detail["ID Konten"].isin(konten_dipilih)].copy()
    if detail.empty:
        raise ValueError("Pilih minimal satu materi, tugas, atau kuis untuk dianalisis.")

    ringkasan_sesi = []
    for (sheet, label), d in detail.groupby(["Sheet", "Sesi"], sort=False):
        ringkasan_sesi.append({
            "Sesi/Sheet": label, "Nama Sheet": sheet,
            "Jumlah Peserta": d["Email"].nunique(),
            "Jumlah Konten": d["ID Konten"].nunique(),
            "Peserta Mengakses": d.loc[d["Diakses"], "Email"].nunique(),
            "Peserta Mengumpulkan Tugas": d.loc[d["Tugas Dikumpulkan"], "Email"].nunique(),
            "Peserta Mengikuti Kuis": d.loc[d["Mengikuti Kuis"], "Email"].nunique(),
        })
    item = detail[["Sheet", "Nomor Item", "Jenis"]].drop_duplicates()
    total_materi = int((item["Jenis"] == "Materi").sum())
    total_tugas = int((item["Jenis"] == "Tugas").sum())
    total_kuis = int((item["Jenis"] == "Kuis").sum())
    hasil = []

    for email, grup in detail.groupby("Email", sort=False):
        nama = grup["Nama"].dropna().iloc[0]
        materi = int(grup.loc[grup["Jenis"] == "Materi", "Diakses"].sum())
        tugas = int(grup["Tugas Dikumpulkan"].sum())
        kuis = int(grup["Mengikuti Kuis"].sum())
        nilai_kuis = grup.loc[grup["Mengikuti Kuis"], "Nilai Kuis"].dropna()
        komponen = []
        if total_materi: komponen.append((materi / total_materi, BOBOT["Materi"]))
        if total_tugas: komponen.append((tugas / total_tugas, BOBOT["Tugas"]))
        if total_kuis: komponen.append((kuis / total_kuis, BOBOT["Kuis"]))
        skor = 100 * sum(n * b for n, b in komponen) / sum(b for _, b in komponen)
        status = "Aktif" if skor >= batas_aktif else "Cukup Aktif" if skor >= batas_cukup else "Perlu Ditindaklanjuti"
        alasan = []
        if total_materi and materi < total_materi: alasan.append(f"belum mengakses {total_materi-materi} materi")
        if total_tugas and tugas < total_tugas: alasan.append(f"belum mengumpulkan {total_tugas-tugas} tugas")
        if total_kuis and kuis < total_kuis: alasan.append(f"belum mengikuti {total_kuis-kuis} kuis")
        hasil.append({
            "Email": email, "Nama": nama, "Materi Diakses": materi, "Total Materi": total_materi,
            "Tugas Dikumpulkan": tugas, "Total Tugas": total_tugas, "Kuis Diikuti": kuis,
            "Total Kuis": total_kuis,
            "Rata-rata Nilai Kuis": round(float(nilai_kuis.mean()), 2) if not nilai_kuis.empty else None,
            "Total Views": int(grup["Views"].sum()), "Total Komentar": int(grup["Komentar"].sum()),
            "Total Unduhan": int(grup["Dokumen Diunduh"].sum()), "Skor Aktivitas": round(skor, 2),
            "Status": status, "Tindak Lanjut": "; ".join(alasan) if alasan else "Tidak ada",
        })

    peserta = pd.DataFrame(hasil).sort_values(["Skor Aktivitas", "Nama"], ascending=[False, True])
    sesi = pd.DataFrame(ringkasan_sesi)
    tindak_lanjut = peserta[peserta["Status"] == "Perlu Ditindaklanjuti"].copy()
    ringkasan = {
        "jumlah_peserta": int(peserta["Email"].nunique()), "jumlah_sheet": len(excel.sheet_names),
        "jumlah_materi": total_materi, "jumlah_tugas": total_tugas, "jumlah_kuis": total_kuis,
        "aktif": int((peserta["Status"] == "Aktif").sum()),
        "cukup_aktif": int((peserta["Status"] == "Cukup Aktif").sum()),
        "perlu_tindak_lanjut": int((peserta["Status"] == "Perlu Ditindaklanjuti").sum()),
    }
    return ringkasan, peserta, sesi, tindak_lanjut, detail


def buat_excel(peserta, sesi, tindak_lanjut, detail):
    keluaran = BytesIO()
    with pd.ExcelWriter(keluaran, engine="openpyxl") as writer:
        for nama, data in [("Rekap Peserta", peserta), ("Ringkasan Sesi", sesi),
                           ("Tindak Lanjut", tindak_lanjut), ("Detail Aktivitas", detail)]:
            data.to_excel(writer, index=False, sheet_name=nama)
        for worksheet in writer.book.worksheets:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            for kolom in worksheet.columns:
                lebar = min(max(len(str(sel.value or "")) for sel in kolom) + 2, 45)
                worksheet.column_dimensions[kolom[0].column_letter].width = lebar
    return keluaran.getvalue()
