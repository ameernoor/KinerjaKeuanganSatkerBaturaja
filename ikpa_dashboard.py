import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
import numpy as np
from pathlib import Path
import os
import base64
from github import Github
from github import Auth

# Konfigurasi halaman
st.set_page_config(
    page_title="Dashboard IKPA KPPN Baturaja",
    page_icon="📊",
    layout="wide"
)

# Inisialisasi session state untuk menyimpan data dan aktivitas
if 'data_storage' not in st.session_state:
    st.session_state.data_storage = {}

if 'activity_log' not in st.session_state:
    st.session_state.activity_log = []  # Each entry: dict with timestamp, action, period, status

# Fungsi untuk memproses file Excel
def process_excel_file(uploaded_file, year):
    """
    Memproses file Excel IKPA sesuai struktur yang telah ditentukan
    """
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        # 1️⃣ Ekstrak bulan dari baris ke-2 (index 1)
        month_text = str(df_raw.iloc[1, 0])
        month = month_text.split(":")[-1].strip() if ":" in month_text else "UNKNOWN"
        
        # 2️⃣ Ekstrak data (baris ke-5 dst)
        df_data = df_raw.iloc[4:].reset_index(drop=True)
        df_data.columns = range(len(df_data.columns))
        
        processed_rows = []
        i = 0
        while i < len(df_data):
            if i + 3 >= len(df_data):
                break
            
            nilai_row = df_data.iloc[i]
            bobot_row = df_data.iloc[i + 1]
            nilai_akhir_row = df_data.iloc[i + 2]
            nilai_aspek_row = df_data.iloc[i + 3]
            
            # Ekstrak kolom
            no = nilai_row[0]
            kode_kppn = str(nilai_row[1]).strip("'") if pd.notna(nilai_row[1]) else ""
            kode_ba = str(nilai_row[2]).strip("'") if pd.notna(nilai_row[2]) else ""
            kode_satker = str(nilai_row[3]).strip("'") if pd.notna(nilai_row[3]) else ""
            uraian_satker = nilai_row[4] if pd.notna(nilai_row[4]) else ""
            
            aspek_perencanaan = nilai_aspek_row[6] if pd.notna(nilai_aspek_row[6]) else 0
            aspek_pelaksanaan = nilai_aspek_row[8] if pd.notna(nilai_aspek_row[8]) else 0
            aspek_hasil = nilai_aspek_row[12] if pd.notna(nilai_aspek_row[12]) else 0
            
            revisi_dipa = nilai_row[6] if pd.notna(nilai_row[6]) else 0
            deviasi_hal3 = nilai_row[7] if pd.notna(nilai_row[7]) else 0
            penyerapan = nilai_row[8] if pd.notna(nilai_row[8]) else 0
            belanja_kontraktual = nilai_row[9] if pd.notna(nilai_row[9]) else 0
            penyelesaian_tagihan = nilai_row[10] if pd.notna(nilai_row[10]) else 0
            pengelolaan_up = nilai_row[11] if pd.notna(nilai_row[11]) else 0
            capaian_output = nilai_row[12] if pd.notna(nilai_row[12]) else 0
            
            nilai_total = nilai_row[13] if pd.notna(nilai_row[13]) else 0
            konversi_bobot = nilai_row[14] if pd.notna(nilai_row[14]) else 0
            dispensasi_spm = nilai_row[15] if pd.notna(nilai_row[15]) else 0
            nilai_akhir = nilai_row[16] if pd.notna(nilai_row[16]) else 0

            # Simpan bobot & nilai terbobot
            bobot_dict = {
                'Revisi DIPA': bobot_row[6], 'Deviasi Halaman III DIPA': bobot_row[7],
                'Penyerapan Anggaran': bobot_row[8], 'Belanja Kontraktual': bobot_row[9],
                'Penyelesaian Tagihan': bobot_row[10], 'Pengelolaan UP dan TUP': bobot_row[11],
                'Capaian Output': bobot_row[12]
            }
            nilai_terbobot_dict = {
                'Revisi DIPA': nilai_akhir_row[6], 'Deviasi Halaman III DIPA': nilai_akhir_row[7],
                'Penyerapan Anggaran': nilai_akhir_row[8], 'Belanja Kontraktual': nilai_akhir_row[9],
                'Penyelesaian Tagihan': nilai_akhir_row[10], 'Pengelolaan UP dan TUP': nilai_akhir_row[11],
                'Capaian Output': nilai_akhir_row[12]
            }

            row_data = {
                'No': no, 'Kode KPPN': kode_kppn, 'Kode BA': kode_ba, 'Kode Satker': kode_satker,
                'Uraian Satker': uraian_satker,
                'Kualitas Perencanaan Anggaran': aspek_perencanaan,
                'Kualitas Pelaksanaan Anggaran': aspek_pelaksanaan,
                'Kualitas Hasil Pelaksanaan Anggaran': aspek_hasil,
                'Revisi DIPA': revisi_dipa, 'Deviasi Halaman III DIPA': deviasi_hal3,
                'Penyerapan Anggaran': penyerapan, 'Belanja Kontraktual': belanja_kontraktual,
                'Penyelesaian Tagihan': penyelesaian_tagihan, 'Pengelolaan UP dan TUP': pengelolaan_up,
                'Capaian Output': capaian_output,
                'Nilai Total': nilai_total, 'Konversi Bobot': konversi_bobot,
                'Dispensasi SPM (Pengurang)': dispensasi_spm,
                'Nilai Akhir (Nilai Total/Konversi Bobot)': nilai_akhir,
                'Bulan': month, 'Tahun': year,
                'Bobot': bobot_dict, 'Nilai Terbobot': nilai_terbobot_dict
            }
            processed_rows.append(row_data)
            i += 4

        df_processed = pd.DataFrame(processed_rows)
        df_processed = df_processed.sort_values('Nilai Akhir (Nilai Total/Konversi Bobot)', ascending=False)
        df_processed['Peringkat'] = range(1, len(df_processed) + 1)
        df_processed['Satker'] = df_processed['Uraian Satker'].astype(str) + ' (' + df_processed['Kode Satker'].astype(str) + ')'
        df_processed['Source'] = 'Upload'
        
        return df_processed, month, year

    except Exception as e:
        st.error(f"Error memproses file: {str(e)}")
        return None, None, None

# Save any file (Excel/template) to your GitHub repo
def save_file_to_github(file_bytes, filename, folder="data"):
    token = st.secrets.get("GITHUB_TOKEN")
    repo_name = st.secrets.get("GITHUB_REPO")

    if not token or not repo_name:
        st.stop()
        st.error("❌ Gagal mengakses GitHub: GITHUB_TOKEN atau GITHUB_REPO tidak ditemukan di secrets.")
        return

    g = Github(auth=Auth.Token(token))
    repo = g.get_repo(repo_name)
    path = f"{folder}/{filename}"

    try:
        contents = repo.get_contents(path)
        repo.update_file(contents.path, f"Update {filename}", file_bytes, contents.sha)
        st.success(f"✅ File {filename} diperbarui di GitHub.")
    except Exception:
        repo.create_file(path, f"Upload {filename}", file_bytes)
        st.success(f"✅ File {filename} diunggah ke GitHub.")


# Load all uploaded data from GitHub (run on startup)
def load_data_from_github():
    token = st.secrets.get("GITHUB_TOKEN")
    repo_name = st.secrets.get("GITHUB_REPO")

    if not token or not repo_name:
        st.stop()
        st.error("❌ Gagal mengakses GitHub: GITHUB_TOKEN atau GITHUB_REPO tidak ditemukan di secrets.")
        return

    g = Github(auth=Auth.Token(token))
    repo = g.get_repo(repo_name)

    try:
        contents = repo.get_contents("data")
    except Exception:
        st.info("📁 Folder 'data' belum ada di repository GitHub.")
        return

    st.session_state.data_storage = {}

    for file in contents:
        if not file.name.endswith(".xlsx"):
            continue

        decoded = base64.b64decode(file.content)
        df = pd.read_excel(io.BytesIO(decoded))
        parts = file.name.replace("IKPA_", "").replace(".xlsx", "").split("_")
        if len(parts) != 2:
            continue
        month, year = parts

        # standardize columns
        if 'Uraian Satker' in df.columns and 'Kode Satker' in df.columns:
            df['Satker'] = df['Uraian Satker'].astype(str) + ' (' + df['Kode Satker'].astype(str) + ')'
        elif 'Kode Satker' in df.columns:
            df['Satker'] = df['Kode Satker'].astype(str)
        else:
            df['Satker'] = df.index.astype(str)

        numeric_cols = [
            'Nilai Akhir (Nilai Total/Konversi Bobot)', 'Nilai Total', 'Konversi Bobot',
            'Revisi DIPA', 'Deviasi Halaman III DIPA', 'Penyerapan Anggaran',
            'Belanja Kontraktual', 'Penyelesaian Tagihan', 'Pengelolaan UP dan TUP', 'Capaian Output'
        ]
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        df['Bulan'] = df.get('Bulan', month)
        df['Tahun'] = df.get('Tahun', year)
        df['Source'] = 'GitHub'
        df['Period'] = f"{month} {year}"
        df['Period_Sort'] = f"{year}-{month}"

        if 'Peringkat' not in df.columns and 'Nilai Akhir (Nilai Total/Konversi Bobot)' in df.columns:
            df = df.sort_values('Nilai Akhir (Nilai Total/Konversi Bobot)', ascending=False)
            df['Peringkat'] = range(1, len(df) + 1)

        st.session_state.data_storage[(str(month), str(year))] = df

    st.success(f"✅ {len(st.session_state.data_storage)} file berhasil dimuat dari GitHub.")

# Path ke file template (akan diatur di session state)
TEMPLATE_PATH = r"C:\Users\KEMENKEU\Desktop\INDIKATOR PELAKSANAAN ANGGARAN.xlsx"

# Fungsi untuk membaca template Excel yang sudah ada
def get_template_file():
    """
    Membaca file template Excel yang sudah ada
    """
    try:
        # Cek apakah file template ada di path default
        if Path(TEMPLATE_PATH).exists():
            with open(TEMPLATE_PATH, 'rb') as f:
                return f.read()
        else:
            # Jika tidak ada, gunakan template dari session state (jika di-upload admin)
            if 'template_file' in st.session_state:
                return st.session_state.template_file
            else:
                return None
    except Exception as e:
        st.error(f"Error membaca template: {str(e)}")
        return None

# Fungsi visualisasi podium/bintang
def create_ranking_chart(df, title, top=True, limit=10):
    """
    Membuat visualisasi ranking dengan bar chart horizontal yang menarik
    (Sekarang menggunakan kolom 'Satker' untuk label agar unik)
    """
    if top:
        df_sorted = df.nlargest(limit, 'Nilai Akhir (Nilai Total/Konversi Bobot)')
        color_scale = 'Greens'
        emoji = '🏆'
    else:
        df_sorted = df.nsmallest(limit, 'Nilai Akhir (Nilai Total/Konversi Bobot)')
        color_scale = 'Reds'
        emoji = '⚠️'
    
    fig = go.Figure()
    
    colors = px.colors.sequential.Greens if top else px.colors.sequential.Reds
    
    # use 'Satker' for y labels to keep them unique
    fig.add_trace(go.Bar(
        y=df_sorted['Satker'],
        x=df_sorted['Nilai Akhir (Nilai Total/Konversi Bobot)'],
        orientation='h',
        marker=dict(
            color=df_sorted['Nilai Akhir (Nilai Total/Konversi Bobot)'],
            colorscale=color_scale,
            showscale=False
        ),
        text=df_sorted['Nilai Akhir (Nilai Total/Konversi Bobot)'].round(2),
        textposition='outside',
        hovertemplate='<b>%{y}</b><br>Nilai: %{x:.2f}<extra></extra>'
    ))
    
    fig.update_layout(
        title=f"{emoji} {title}",
        xaxis_title="Nilai Akhir",
        yaxis_title="",
        height=max(400, limit * 40),
        yaxis={'categoryorder': 'total ascending' if not top else 'total descending'},
        showlegend=False
    )
    
    return fig

# Fungsi visualisasi untuk satker bermasalah
def create_problem_chart(df, column, threshold, title, comparison='less'):
    """
    Membuat visualisasi untuk satker dengan masalah
    (Sekarang menggunakan kolom 'Satker' untuk label agar unik)
    """
    if comparison == 'less':
        df_filtered = df[df[column] < threshold]
    else:
        df_filtered = df[df[column] > threshold]
    
    if len(df_filtered) == 0:
        return None
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        y=df_filtered['Satker'],
        x=df_filtered[column],
        orientation='h',
        marker=dict(
            color=df_filtered[column],
            colorscale='RdYlGn',
            showscale=True,
            cmin=0,
            cmax=100
        ),
        text=df_filtered[column].round(2),
        textposition='outside',
        hovertemplate='<b>%{y}</b><br>Nilai: %{x:.2f}<extra></extra>'
    ))
    
    fig.add_vline(x=threshold, line_dash="dash", line_color="red", 
                  annotation_text=f"Target: {threshold}")
    
    fig.update_layout(
        title=f"⚠️ {title}",
        xaxis_title="Nilai",
        yaxis_title="",
        height=max(400, len(df_filtered) * 40),
        yaxis={'categoryorder': 'total ascending'},
        showlegend=False
    )
    
    return fig

# HALAMAN 1: DASHBOARD UTAMA
def page_dashboard():
    st.title("📊 Dashboard Utama IKPA Satker Mitra KPPN Baturaja")
    
    if not st.session_state.data_storage:
        st.warning("⚠️ Belum ada data yang diunggah. Silakan unggah data melalui halaman Admin.")
        return
    
    # Dapatkan data terbaru
    all_periods = sorted(st.session_state.data_storage.keys(), reverse=True)
    
    if not all_periods:
        st.warning("⚠️ Belum ada data yang tersedia.")
        return
    
    # Filter periode
    col1, col2 = st.columns([3, 1])
    with col1:
        selected_period = st.selectbox(
            "📅 Pilih Periode",
            options=all_periods,
            index=0,
            format_func=lambda x: f"{x[0]} {x[1]}"
        )
    
    df = st.session_state.data_storage[selected_period]
    
    st.markdown("---")
    
    # Metrik ringkasan
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📋 Total Satker", len(df))
    with col2:
        avg_score = df['Nilai Akhir (Nilai Total/Konversi Bobot)'].mean()
        st.metric("📈 Rata-rata Nilai", f"{avg_score:.2f}")
    with col3:
        perfect_count = len(df[df['Nilai Akhir (Nilai Total/Konversi Bobot)'] == 100])
        st.metric("⭐ Satker Nilai 100", perfect_count)
    with col4:
        below_80 = len(df[df['Nilai Akhir (Nilai Total/Konversi Bobot)'] < 80])
        st.metric("⚠️ Satker < 80", below_80)
    
    st.markdown("---")
    
    # Top 10 dengan Belanja Kontraktual
    st.subheader("🏆 Top 10 Satker dengan Belanja Kontraktual")
    df_with_kontrak = df[df['Belanja Kontraktual'] != 0]
    if len(df_with_kontrak) > 0:
        fig1 = create_ranking_chart(df_with_kontrak, "Satker Terbaik (Dengan Belanja Kontraktual)", top=True, limit=10)
        st.plotly_chart(fig1, use_container_width=True)
    else:
        st.info("Tidak ada satker dengan belanja kontraktual.")
    
    # Top 10 tanpa Belanja Kontraktual
    st.subheader("🏆 Top 10 Satker tanpa Belanja Kontraktual")
    df_without_kontrak = df[df['Belanja Kontraktual'] == 0]
    if len(df_without_kontrak) > 0:
        fig2 = create_ranking_chart(df_without_kontrak, "Satker Terbaik (Tanpa Belanja Kontraktual)", top=True, limit=10)
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Tidak ada satker tanpa belanja kontraktual.")
    
    st.markdown("---")
    
    # Bottom 10 dengan Belanja Kontraktual
    st.subheader("📉 Bottom 10 Satker dengan Belanja Kontraktual")
    if len(df_with_kontrak) > 0:
        fig3 = create_ranking_chart(df_with_kontrak, "Satker Perlu Perbaikan (Dengan Belanja Kontraktual)", top=False, limit=10)
        st.plotly_chart(fig3, use_container_width=True)
    
    # Bottom 10 tanpa Belanja Kontraktual
    st.subheader("📉 Bottom 10 Satker tanpa Belanja Kontraktual")
    if len(df_without_kontrak) > 0:
        fig4 = create_ranking_chart(df_without_kontrak, "Satker Perlu Perbaikan (Tanpa Belanja Kontraktual)", top=False, limit=10)
        st.plotly_chart(fig4, use_container_width=True)
    
    st.markdown("---")
    
    # Satker dengan masalah
    st.subheader("🚨 Satker yang Memerlukan Perhatian Khusus")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Deviasi Hal 3 DIPA
        fig_dev = create_problem_chart(
            df, 
            'Deviasi Halaman III DIPA', 
            90, 
            "Deviasi Hal 3 DIPA Belum Optimal (< 90)",
            'less'
        )
        if fig_dev:
            st.plotly_chart(fig_dev, use_container_width=True)
        else:
            st.success("✅ Semua satker sudah optimal untuk Deviasi Hal 3 DIPA")
    
    with col2:
        # Pengelolaan UP dan TUP
        fig_up = create_problem_chart(
            df, 
            'Pengelolaan UP dan TUP', 
            100, 
            "Pengelolaan UP dan TUP Belum Optimal (< 100)",
            'less'
        )
        if fig_up:
            st.plotly_chart(fig_up, use_container_width=True)
        else:
            st.success("✅ Semua satker sudah optimal untuk Pengelolaan UP dan TUP")
    
    # Capaian Output
    fig_output = create_problem_chart(
        df, 
        'Capaian Output', 
        100, 
        "Capaian Output Belum Optimal (< 100)",
        'less'
    )
    if fig_output:
        st.plotly_chart(fig_output, use_container_width=True)
    else:
        st.success("✅ Semua satker sudah optimal untuk Capaian Output")
    
    st.markdown("---")
    
    # Tabel detail
    st.subheader("📋 Tabel Detail Satker")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        view_mode = st.radio(
            "Tampilan",
            options=['aspek', 'komponen'],
            format_func=lambda x: 'Berdasarkan Aspek' if x == 'aspek' else 'Berdasarkan Komponen',
            horizontal=True
        )
    
    with col2:
        if view_mode == 'komponen':
            value_type = st.selectbox(
                "Jenis Nilai",
                options=['Nilai', 'Bobot', 'Nilai Terbobot']
            )
    
    # Kolom yang ditampilkan
    display_columns = ['Peringkat', 'Kode BA', 'Kode Satker', 'Uraian Satker']
    
    if view_mode == 'aspek':
        display_columns += [
            'Kualitas Perencanaan Anggaran',
            'Kualitas Pelaksanaan Anggaran',
            'Kualitas Hasil Pelaksanaan Anggaran'
        ]
        df_display = df[display_columns + ['Nilai Total', 'Konversi Bobot', 
                                            'Dispensasi SPM (Pengurang)', 
                                            'Nilai Akhir (Nilai Total/Konversi Bobot)']].copy()
    else:
        component_cols = [
            'Revisi DIPA', 'Deviasi Halaman III DIPA', 'Penyerapan Anggaran',
            'Belanja Kontraktual', 'Penyelesaian Tagihan', 
            'Pengelolaan UP dan TUP', 'Capaian Output'
        ]
        
        df_display = df[display_columns + ['Nilai Total', 'Konversi Bobot', 
                                            'Dispensasi SPM (Pengurang)', 
                                            'Nilai Akhir (Nilai Total/Konversi Bobot)']].copy()
        
        if value_type == 'Nilai':
            for col in component_cols:
                df_display[col] = df[col]
        elif value_type == 'Bobot':
            for col in component_cols:
                df_display[col] = df['Bobot'].apply(lambda x: x.get(col, 0) if isinstance(x, dict) else 0)
        else:  # Nilai Terbobot
            for col in component_cols:
                df_display[col] = df['Nilai Terbobot'].apply(lambda x: x.get(col, 0) if isinstance(x, dict) else 0)
        
        # Reorder kolom
        final_cols = display_columns + component_cols + ['Nilai Total', 'Konversi Bobot', 
                                                          'Dispensasi SPM (Pengurang)', 
                                                          'Nilai Akhir (Nilai Total/Konversi Bobot)']
        df_display = df_display[final_cols]
    
    # Styling untuk tabel
    def highlight_top(s):
        if s.name == 'Peringkat':
            return ['background-color: gold' if v <= 3 else '' for v in s]
        return ['' for _ in s]
    
    st.dataframe(
        df_display.style.apply(highlight_top).format(precision=2),
        use_container_width=True,
        height=600
    )

# HALAMAN 2: DASHBOARD TREN HISTORIS
def page_trend():
    st.title("📈 Dashboard Tren Historis IKPA Satker Mitra Kerja KPPN Baturaja")
    
    if not st.session_state.data_storage:
        st.warning("⚠️ Belum ada data yang diunggah. Silakan unggah data melalui halaman Admin.")
        return
    
    # Gabungkan semua data
    all_data = []
    for period, df in st.session_state.data_storage.items():
        df_copy = df.copy()
        # ensure Period & Period_Sort exist
        df_copy['Period'] = f"{period[0]} {period[1]}"
        df_copy['Period_Sort'] = f"{period[1]}-{period[0]}"
        all_data.append(df_copy)
    
    if not all_data:
        st.warning("⚠️ Belum ada data historis yang tersedia.")
        return
    
    df_all = pd.concat(all_data, ignore_index=True)
    
    # Filter periode
    st.subheader("🎯 Filter Analisis")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        available_periods = sorted(df_all['Period_Sort'].unique())
        start_period = st.selectbox(
            "Periode Awal",
            options=available_periods,
            index=0
        )
    
    with col2:
        end_period = st.selectbox(
            "Periode Akhir",
            options=available_periods,
            index=len(available_periods) - 1
        )
    
    # Filter berdasarkan periode
    df_filtered = df_all[
        (df_all['Period_Sort'] >= start_period) & 
        (df_all['Period_Sort'] <= end_period)
    ]
    
    with col3:
        # Pilihan metrik
        metric_options = {
            'Nilai Akhir (Nilai Total/Konversi Bobot)': 'Nilai Akhir (Nilai Total/Konversi Bobot)',
            'Kualitas Perencanaan Anggaran': 'Kualitas Perencanaan Anggaran',
            'Kualitas Pelaksanaan Anggaran': 'Kualitas Pelaksanaan Anggaran',
            'Kualitas Hasil Pelaksanaan Anggaran': 'Kualitas Hasil Pelaksanaan Anggaran',
            'Revisi DIPA': 'Revisi DIPA',
            'Deviasi Halaman III DIPA': 'Deviasi Halaman III DIPA',
            'Penyerapan Anggaran': 'Penyerapan Anggaran',
            'Belanja Kontraktual': 'Belanja Kontraktual',
            'Penyelesaian Tagihan': 'Penyelesaian Tagihan',
            'Pengelolaan UP dan TUP': 'Pengelolaan UP dan TUP',
            'Capaian Output': 'Capaian Output'
        }
        
        selected_metric = st.selectbox(
            "Metrik yang Ditampilkan",
            options=list(metric_options.keys()),
            index=0
        )
    
    # Pilih satker
    # All keys are (month_str, year_str). To sort by year then month, create sortable key:
    def period_sort_key(k):
        mon, yr = k
        # convert year to int if possible, month remain string but sorting will be stable for same year
        try:
            y = int(yr)
        except:
            y = yr
        return (y, mon)

    latest_period = sorted(st.session_state.data_storage.keys(), key=period_sort_key, reverse=True)[0]
    latest_df = st.session_state.data_storage[latest_period].copy()
    # Make sure 'Kode Satker' exists and is a string
    if 'Kode Satker' in latest_df.columns:
        latest_df['Kode Satker'] = latest_df['Kode Satker'].astype(str)
    else:
        latest_df['Kode Satker'] = latest_df.index.astype(str)

    bottom_10_default = latest_df.nsmallest(10, 'Nilai Akhir (Nilai Total/Konversi Bobot)')['Kode Satker'].astype(str).tolist()
    
    # use the new 'Satker' column for selection (unique)
    all_satker = sorted(df_all['Satker'].unique())
    selected_satker = st.multiselect(
    "Pilih Satker",
    options=all_satker,
    default=[s for s in all_satker if any(str(code) in s for code in bottom_10_default)][:10]
    )
    
    if not selected_satker:
        st.warning("Silakan pilih minimal satu satker untuk melihat tren.")
        return
    
    # Filter berdasarkan satker (use 'Satker' to avoid duplicate names)
    df_plot = df_filtered[df_filtered['Satker'].isin(selected_satker)]
    
    # Buat line chart
    fig = go.Figure()
    for satker in selected_satker:
        df_satker = df_plot[df_plot['Satker'] == satker].sort_values('Period_Sort')
        
        fig.add_trace(go.Scatter(
            x=df_satker['Period'],
            y=df_satker[selected_metric],
            mode='lines+markers',
            name=satker,
            hovertemplate='<b>%{fullData.name}</b><br>Periode: %{x}<br>Nilai: %{y:.2f}<extra></extra>'
        ))
    
    fig.update_layout(
        title=f"Tren {selected_metric}",
        xaxis_title="Periode",
        yaxis_title="Nilai",
        height=600,
        hovermode='x unified',
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02
        )
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Analisis tren
    st.markdown("---")
    st.subheader("⚠️ Analisis Tren dan Peringatan")
    
    warnings = []
    
    for satker in selected_satker:
        df_satker = df_plot[df_plot['Satker'] == satker].sort_values('Period_Sort')
        
        if len(df_satker) >= 2:
            values = df_satker[selected_metric].values
            
            # Cek tren menurun (2 periode terakhir)
            if len(values) >= 2:
                last_value = values[-1]
                prev_value = values[-2]
                
                if last_value < prev_value:
                    decrease = prev_value - last_value
                    warnings.append({
                        'Satker': satker,
                        'Metrik': selected_metric,
                        'Nilai Sebelumnya': prev_value,
                        'Nilai Terkini': last_value,
                        'Penurunan': decrease
                    })
    
    if warnings:
        st.warning(f"⚠️ Ditemukan {len(warnings)} satker dengan tren menurun!")
        
        for w in warnings:
            st.markdown(f"""
            **{w['Satker']}**  
            - Metrik: {w['Metrik']}
            - Nilai sebelumnya: {w['Nilai Sebelumnya']:.2f}
            - Nilai terkini: {w['Nilai Terkini']:.2f}
            - Penurunan: {w['Penurunan']:.2f} poin
            """)
            st.markdown("---")
    else:
        st.success("✅ Tidak ada satker dengan tren menurun pada periode yang dipilih!")

# HALAMAN 3: ADMIN
def page_admin():
    st.title("🔐 Halaman Administrasi")
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        st.warning("🔒 Halaman ini memerlukan autentikasi")
        password = st.text_input("Masukkan Password", type="password")
        if st.button("Login"):
            if password == "109KPPN":
                st.session_state.authenticated = True
                st.success("✅ Login berhasil!")
                st.rerun()
            else:
                st.error("❌ Password salah!")
        return

    st.success("✅ Anda telah login sebagai Admin")
    # 🧩 Optional: Debug GitHub connection
    with st.expander("🧩 Debug GitHub Connection"):
        try:
            token = st.secrets["GITHUB_TOKEN"]
            repo_name = st.secrets["GITHUB_REPO"]
            g = Github(auth=Auth.Token(token))
            repo = g.get_repo(repo_name)
            st.success(f"Terhubung ke GitHub repo: {repo.full_name}")
        except Exception as e:
            st.error(f"❌ Gagal terhubung ke GitHub: {e}")
    
    if st.button("🚪 Logout"):
        st.session_state.authenticated = False
        st.rerun()

    st.markdown("---")
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📤 Upload Data", 
        "🗑️ Hapus Data", 
        "📥 Download Data", 
        "📋 Download Template", 
        "🕓 Riwayat Aktivitas"
    ])

    # TAB 1: UPLOAD DATA
    with tab1:
        st.subheader("📤 Upload Data Bulanan")

        upload_year = st.selectbox(
            "Pilih Tahun",
            list(range(2020, 2031)),
            index=list(range(2020, 2031)).index(datetime.now().year)
        )

        uploaded_file = st.file_uploader("Pilih file Excel IKPA", type=['xlsx', 'xls'])

        if uploaded_file is not None:
            if st.button("🔄 Proses Data", type="primary"):
                with st.spinner("Memproses data..."):
                    df_processed, month, year = process_excel_file(uploaded_file, upload_year)

                    if df_processed is None:
                        st.error("❌ Gagal memproses file. Pastikan format file sesuai template.")
                        st.stop()

                    period_key = (str(month), str(year))
                    filename = f"IKPA_{month}_{year}.xlsx"

                    # ✅ Ask confirmation if data already exists
                    if period_key in st.session_state.data_storage:
                        st.warning(f"⚠️ Data untuk **{month} {year}** sudah ada di sistem dan GitHub.")
                        confirm_replace = st.checkbox("Saya yakin ingin mengganti data yang sudah ada.", key=f"confirm_replace_{month}_{year}")

                        if not confirm_replace:
                            st.info("🕒 Unggahan dibatalkan sampai Anda mencentang konfirmasi.")
                            st.stop()

                    # ✅ Proceed upload
                    try:
                        df_processed['Kode Satker'] = df_processed['Kode Satker'].astype(str)
                        st.session_state.data_storage[period_key] = df_processed

                        excel_bytes = io.BytesIO()
                        with pd.ExcelWriter(excel_bytes, engine='openpyxl') as writer:
                            df_excel = df_processed.drop(['Bobot', 'Nilai Terbobot'], axis=1, errors='ignore')
                            df_excel.to_excel(writer, index=False, sheet_name='Data IKPA')
                        excel_bytes.seek(0)

                        save_file_to_github(excel_bytes.getvalue(), filename, folder="data")

                        # ✅ Professional success feedback
                        st.toast(f"✅ Data {month} {year} berhasil diunggah & disimpan di GitHub.", icon="✅")
                        st.success(f"✅ Data {month} {year} tersimpan dengan aman di sistem dan GitHub.")
                        st.info("💾 Data berhasil diperbarui. Anda dapat melihatnya di halaman Dashboard Utama.")
                        st.snow()

                        st.session_state.activity_log.append({
                            "Waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Aksi": "Upload/Replace",
                            "Periode": f"{month} {year}",
                            "Status": "✅ Sukses disimpan ke GitHub"
                        })
                    except Exception as e:
                        st.error(f"❌ Gagal menyimpan ke GitHub: {e}")
        
    # TAB 2: HAPUS DATA
    with tab2:
        st.subheader("🗑️ Hapus Data Bulanan")

        if not st.session_state.data_storage:
            st.info("ℹ️ Belum ada data yang tersimpan.")
        else:
            available_periods = sorted(st.session_state.data_storage.keys(), reverse=True)

            period_to_delete = st.selectbox(
                "Pilih periode yang akan dihapus",
                options=available_periods,
                format_func=lambda x: f"{x[0]} {x[1]}"
            )

            month, year = period_to_delete
            filename = f"data/IKPA_{month}_{year}.xlsx"

            confirm_delete = st.checkbox(
                f"⚠️ Saya yakin ingin menghapus data {month} {year} dari sistem dan GitHub.",
                key=f"confirm_delete_{month}_{year}"
            )

            if st.button("🗑️ Hapus Data Ini", type="primary"):
                if not confirm_delete:
                    st.info("🕒 Penghapusan dibatalkan sampai Anda mencentang konfirmasi.")
                    st.stop()

                # 1️⃣ Hapus dari session_state
                del st.session_state.data_storage[period_to_delete]
                st.success(f"✅ Data {month} {year} berhasil dihapus dari session.")

                # 2️⃣ Hapus juga dari GitHub
                try:
                    token = st.secrets.get("GITHUB_TOKEN")
                    repo_name = st.secrets.get("GITHUB_REPO")

                    if not token or not repo_name:
                        raise ValueError("GitHub credentials tidak ditemukan di secrets.")

                    g = Github(auth=Auth.Token(token))
                    repo = g.get_repo(repo_name)
                    full_path = f"data/IKPA_{month}_{year}.xlsx"

                    contents = repo.get_contents(full_path)

                    # Handle both file and list-of-files cases
                    if isinstance(contents, list):
                        target = next((c for c in contents if c.path == full_path), None)
                        if not target:
                            raise FileNotFoundError(f"File {full_path} tidak ditemukan di GitHub.")
                        contents = target

                    repo.delete_file(
                        contents.path,
                        message=f"🗑️ delete {full_path}",
                        sha=contents.sha
                    )

                    st.toast(f"🗑️ File {full_path} berhasil dihapus dari GitHub.", icon="✅")
                    st.success(f"✅ File {month} {year} dihapus dari GitHub dan lokal.")
                    st.snow()
                    st.session_state.activity_log.append({
                        "Waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Aksi": "Hapus",
                        "Periode": f"{month} {year}",
                        "Status": "🗑️ Dihapus dari GitHub dan lokal"
                    })
                    st.info("🔁 Memuat ulang halaman untuk memperbarui tampilan data...")
                    st.rerun()                    

                except Exception as e:
                    st.error(f"❌ Gagal menghapus file dari GitHub untuk {month} {year}.")
                    with st.expander("⚠️ Rincian Error Penghapusan"):
                        st.exception(e)

                    st.session_state.activity_log.append({
                        "Waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Aksi": "Hapus",
                        "Periode": f"{month} {year}",
                        "Status": f"⚠️ Gagal menghapus: {e}"
                    })
                    st.stop()  # ❌ Stop instead of rerun, so error stays visible

    # TAB 3: DOWNLOAD DATA
    with tab3:
        st.subheader("📥 Download Data Bersih")

        if not st.session_state.data_storage:
            st.info("ℹ️ Belum ada data yang tersimpan.")
        else:
            available_periods = sorted(st.session_state.data_storage.keys(), reverse=True)

            period_to_download = st.selectbox(
                "Pilih periode yang akan didownload",
                options=available_periods,
                format_func=lambda x: f"{x[0]} {x[1]}"
            )

            df_download = st.session_state.data_storage[period_to_download]

            # Preview
            with st.expander("👁️ Preview Data"):
                st.dataframe(df_download.head(10), use_container_width=True)

            # Prepare download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Drop only if present; ignore if not present
                df_excel = df_download.drop(['Bobot', 'Nilai Terbobot'], axis=1, errors='ignore')
                # If df_excel is empty, create a minimal sheet to avoid openpyxl "no visible sheet" error
                if df_excel.shape[0] == 0 or df_excel.shape[1] == 0:
                    # create a minimal DataFrame with a placeholder column
                    df_excel = pd.DataFrame({'No data': []})
                df_excel.to_excel(writer, index=False, sheet_name='Data IKPA')

            output.seek(0)

            st.download_button(
                label="📥 Download Excel",
                data=output,
                file_name=f"IKPA_{period_to_download[0]}_{period_to_download[1]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # TAB 4: DOWNLOAD TEMPLATE
    with tab4:
        st.subheader("📋 Download Template Excel")

        st.info("""
        📝 **Panduan Penggunaan Template:**
        1. Template ini adalah contoh format Excel yang harus diikuti.
        2. Baris 1: Judul "INDIKATOR PELAKSANAAN ANGGARAN"
        3. Baris 2: "Sampai Dengan : [NAMA BULAN]"
        4. Baris 3-4: Header kolom (sudah terformat)
        5. Baris 5 dst: Data satker (4 baris per satker)
        ⚠️ Pastikan format angka konsisten (gunakan titik untuk desimal)
        """)

        # Ambil template dari GitHub (folder templates/)
        try:
            token = st.secrets["GITHUB_TOKEN"]
            repo_name = st.secrets["GITHUB_REPO"]
            g = Github(auth=Auth.Token(token))
            repo = g.get_repo(repo_name)
            file_content = repo.get_contents("templates/Template_IKPA.xlsx")
            template_data = base64.b64decode(file_content.content)
        except Exception:
            template_data = get_template_file()

        if template_data:
            st.download_button(
                label="📥 Download Template Excel (dari GitHub atau lokal)",
                data=template_data,
                file_name="Template_IKPA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ Template belum ditemukan di GitHub maupun lokal.")

        st.markdown("---")
        st.subheader("📤 Upload Template Baru")
        st.info("Jika Anda ingin memperbarui template di GitHub, upload di sini:")

        template_upload = st.file_uploader(
            "Upload file template IKPA",
            type=['xlsx', 'xls'],
            key="template_uploader"
        )

        if template_upload:
            try:
                save_file_to_github(template_upload.read(), "Template_IKPA.xlsx", folder="templates")
                st.success("✅ Template berhasil di-upload dan disimpan di GitHub!")
            except Exception as e:
                st.error(f"❌ Gagal upload template ke GitHub: {e}")

        st.markdown("---")
        st.subheader("📊 Informasi Data Tersimpan")

        if st.session_state.data_storage:
            summary_data = []
            for period, df in sorted(st.session_state.data_storage.items(), reverse=True):
                summary_data.append({
                    'Bulan': period[0],
                    'Tahun': period[1],
                    'Jumlah Satker': len(df),
                    'Rata-rata Nilai Akhir': df['Nilai Akhir (Nilai Total/Konversi Bobot)'].mean()
                })

            df_summary = pd.DataFrame(summary_data)
            st.dataframe(df_summary.style.format({'Rata-rata Nilai Akhir': '{:.2f}'}), use_container_width=True)
        else:
            st.info("ℹ️ Belum ada data yang tersimpan.")

    # TAB 5: RIWAYAT AKTIVITAS
    with tab5:
        st.subheader("🕓 Log Aktivitas GitHub (Upload / Delete)")
        st.info("Log ini merekam setiap unggahan, penggantian, dan penghapusan data ke/dari GitHub.")

        if not st.session_state.activity_log:
            st.warning("Belum ada aktivitas tercatat.")
        else:
            df_log = pd.DataFrame(st.session_state.activity_log)
            df_log = df_log[::-1].reset_index(drop=True)  # show latest first
            st.dataframe(
                df_log.style.format(na_rep="-"),
                use_container_width=True,
                height=400
            )

            col1, col2 = st.columns([1, 2])
            with col1:
                if st.button("🧹 Bersihkan Log"):
                    st.session_state.activity_log = []
                    st.success("🧹 Log aktivitas telah dibersihkan.")
                    st.rerun()
            with col2:
                st.download_button(
                    label="📥 Unduh Log Aktivitas (Excel)",
                    data=df_log.to_csv(index=False).encode('utf-8'),
                    file_name="Log_Aktivitas_GitHub.csv",
                    mime="text/csv"
                )

# ===============================
# 🔹 MAIN APP
# ===============================
def main():
    # ✅ Load data dari GitHub jika session_state kosong (hanya sekali di awal)
    if not st.session_state.get("data_storage"):
        with st.spinner("🔄 Memuat data dari GitHub..."):
            try:
                load_data_from_github()
            except Exception as e:
                st.error(f"⚠️ Gagal memuat data dari GitHub: {e}")
    
    # ===============================
    # 🔹 Sidebar Navigation
    # ===============================
    st.sidebar.title("🧭 Navigasi")
    st.sidebar.markdown("---")

    page = st.sidebar.radio(
        "Pilih Halaman",
        options=[
            "📊 Dashboard Utama",
            "📈 Tren Historis",
            "🔐 Admin"
        ],
        index=0
    )

    st.sidebar.markdown("---")
    st.sidebar.info("""
    **Dashboard IKPA**  
    Indikator Kinerja Pelaksanaan Anggaran  
    KPPN Baturaja

    📧 Support: ameer.noor@kemenkeu.go.id
    """)

    # ===============================
    # 🔹 Routing Halaman
    # ===============================
    if page == "📊 Dashboard Utama":
        try:
            page_dashboard()
        except Exception as e:
            st.error(f"❌ Terjadi kesalahan di Dashboard Utama: {e}")

    elif page == "📈 Tren Historis":
        try:
            page_trend()
        except Exception as e:
            st.error(f"❌ Terjadi kesalahan di Dashboard Tren Historis: {e}")

    elif page == "🔐 Admin":
        try:
            page_admin()
        except Exception as e:
            st.error(f"❌ Terjadi kesalahan di Halaman Admin: {e}")

# ===============================
# 🔹 ENTRY POINT
# ===============================
if __name__ == "__main__":
    main()
