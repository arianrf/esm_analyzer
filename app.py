"""Antarmuka ESM Activity Analyzer berbasis Streamlit."""

import streamlit as st
from core import analisis_edlink, buat_excel, daftar_konten

st.set_page_config(page_title="ESM Activity Analyzer", page_icon="logo.png", layout="wide")
kolom_logo, kolom_judul = st.columns([1, 6])

with kolom_logo:
    st.image("logo.png", width=110)
st.title("ESM Activity Analyzer")
st.write("Analisis sederhana terhadap file Excel laporan aktivitas yang diunduh dari LMS Edlink.")

with st.expander("Cara menggunakan"):
    st.write("Unduh laporan aktivitas dari menu Laporan pada Edlink, lalu unggah file Excel di bawah ini. Program membaca seluruh sheet secara otomatis.")
    st.caption("Program tidak membutuhkan akun atau kata sandi Edlink.")

with st.expander("Penyusun"):
    st.markdown("""
    - Muhammad Ammar Naufal, S.Pd., M.Ed., Ph.D.
    - Dr. Syahrullah Asyari, S.Pd., M.Pd.
    - Arian Nurrifahi, S.Kom., M.T.
    - Fettyana, S.Kom., M.Kom.
    - Ahmad Zuhudy Bahtiar, M.Pd.
    - Ibnu Hajar
    - Fifi Alya Sa'dliani
    """)

file = st.file_uploader("Unggah laporan Excel Edlink", type=["xlsx", "xls"])
if file is not None:
    try:
        data_file = file.getvalue()
        konten = daftar_konten(data_file)
        label_konten = {
            row["ID Konten"]: f'{row["Sesi"]} | {row["Jenis"]} | {row["Judul"]}'
            for _, row in konten.iterrows()
        }

        st.subheader("Pengaturan Analisis")
        st.info("Pilih hanya kegiatan yang sudah diberikan atau sudah jatuh tempo.")
        pilihan = st.multiselect(
            "Materi, tugas, dan kuis yang dinilai",
            options=list(label_konten),
            default=list(label_konten),
            format_func=lambda kode: label_konten[kode],
        )
        pengaturan1, pengaturan2 = st.columns(2)
        batas_cukup = pengaturan1.number_input(
            "Batas minimal Cukup Aktif", min_value=0, max_value=99, value=50, step=5
        )
        batas_aktif = pengaturan2.number_input(
            "Batas minimal Aktif", min_value=1, max_value=100, value=75, step=5
        )
        if batas_aktif <= batas_cukup:
            st.error("Batas Aktif harus lebih tinggi daripada batas Cukup Aktif.")
            st.stop()

        ringkasan, peserta, sesi, tindak_lanjut, detail = analisis_edlink(
            data_file, pilihan, batas_aktif, batas_cukup
        )
        st.subheader("Ringkasan Kelas")
        kolom = st.columns(4)
        kolom[0].metric("Jumlah Peserta", ringkasan["jumlah_peserta"])
        kolom[1].metric("Aktif", ringkasan["aktif"])
        kolom[2].metric("Cukup Aktif", ringkasan["cukup_aktif"])
        kolom[3].metric("Perlu Ditindaklanjuti", ringkasan["perlu_tindak_lanjut"])
        st.caption(f"Dianalisis {len(pilihan)} kegiatan terpilih: {ringkasan['jumlah_materi']} materi, {ringkasan['jumlah_tugas']} tugas, dan {ringkasan['jumlah_kuis']} kuis.")

        tab1, tab2, tab3 = st.tabs(["Rekap Peserta", "Ringkasan Sesi", "Tindak Lanjut"])
        with tab1:
            st.dataframe(peserta, use_container_width=True, hide_index=True)
        with tab2:
            st.dataframe(sesi, use_container_width=True, hide_index=True)
        with tab3:
            if tindak_lanjut.empty:
                st.success("Tidak ada peserta dalam kategori Perlu Ditindaklanjuti.")
            else:
                st.dataframe(tindak_lanjut, use_container_width=True, hide_index=True)

        distribusi = peserta["Status"].value_counts().rename_axis("Status").to_frame("Jumlah")
        st.subheader("Distribusi Status Aktivitas")
        st.bar_chart(distribusi)
        st.download_button(
            "Unduh Hasil Analisis Excel", buat_excel(peserta, sesi, tindak_lanjut, detail),
            "hasil_analisis_edlink.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as kesalahan:
        st.error(f"File tidak dapat dianalisis: {kesalahan}")
