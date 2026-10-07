import streamlit as st
import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta
import io

# --- Konfigurasi Halaman ---
st.set_page_config(page_title="Pampam DSCR", layout="wide")
st.title("Pampam DSCR - Simulasi Repayment Capacity & Proyeksi SLIK")
st.markdown("Alat ukur kelayakan kredit konsolidasi dengan Pemisahan Pokok & Bunga (Threshold DSCR > 1.2x)")

# Inisialisasi Session State untuk menampung data excel agar tidak hilang saat direfresh
if 'hasil_excel' not in st.session_state:
    st.session_state['hasil_excel'] = None
if 'sudah_dihitung' not in st.session_state:
    st.session_state['sudah_dihitung'] = False

# --- 1. Upload File SLIK ---
uploaded_file = st.file_uploader("Upload File Rekap SLIK (Excel)", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        
        current_debitur = None
        debitur_list = []
        for index, row in df.iterrows():
            if pd.isna(row['Status Fasilitas']):
                bank_val = str(row['Bank'])
                if "Subtotal" not in bank_val and "GRAND TOTAL" not in bank_val and bank_val != 'nan':
                    current_debitur = bank_val.split(" - No. Laporan")[0].strip()
            debitur_list.append(current_debitur)
        
        df['Nama Debitur'] = debitur_list
        df_aktif = df[df['Status Fasilitas'] == 'Fasilitas Aktif'].copy()
        
        if df_aktif.empty:
            st.warning("Tidak ada fasilitas aktif yang ditemukan dalam file SLIK ini.")
        else:
            st.subheader("1. Pilih Entitas / Pengurus")
            selected_debiturs = st.multiselect(
                "Pilih nama yang kewajibannya akan dikonsolidasikan menjadi beban:", 
                df_aktif['Nama Debitur'].dropna().unique()
            )
            
            if selected_debiturs:
                df_selected = df_aktif[df_aktif['Nama Debitur'].isin(selected_debiturs)].copy()
                
                df_selected = df_selected[['Nama Debitur', 'Bank', 'Fasilitas', 'Limit', 'Baki Debet', 'Suku Bunga', 'Awal Kredit', 'Jatuh Tempo']]
                df_selected.insert(0, 'Dihitung?', True)
                df_selected['Jenis (Rev/Non-Rev)'] = 'Non-Rev'
                
                df_selected['Limit_Str'] = df_selected['Limit'].apply(lambda x: f"Rp {x:,.0f}" if pd.notna(x) else "Rp 0")
                df_selected['BakiDebet_Str'] = df_selected['Baki Debet'].apply(lambda x: f"Rp {x:,.0f}" if pd.notna(x) else "Rp 0")
                
                df_display = df_selected[['Dihitung?', 'Nama Debitur', 'Bank', 'Fasilitas', 'Limit_Str', 'BakiDebet_Str', 'Suku Bunga', 'Awal Kredit', 'Jatuh Tempo', 'Jenis (Rev/Non-Rev)']]
                
                st.subheader("2. Penyesuaian Fasilitas Eksisting")
                st.write("Uncheck 'Dihitung?' untuk fasilitas yang tidak relevan. Tentukan Rev/Non-Rev pada kolom paling kanan.")
                
                edited_df = st.data_editor(
                    df_display,
                    column_config={
                        "Dihitung?": st.column_config.CheckboxColumn("Dihitung?", default=True),
                        "Limit_Str": st.column_config.TextColumn("Limit", disabled=True),
                        "BakiDebet_Str": st.column_config.TextColumn("Baki Debet", disabled=True),
                        "Suku Bunga": st.column_config.NumberColumn("Bunga (%)", format="%.2f %%", disabled=True),
                        "Jenis (Rev/Non-Rev)": st.column_config.SelectboxColumn("Jenis Fasilitas", options=["Rev", "Non-Rev"], required=True)
                    },
                    hide_index=True,
                    use_container_width=True
                )
                
                # --- 3. Input Finansial & Rencana Kredit ---
                st.subheader("3. Input Finansial & Rencana Kredit Baru")
                col1, col2 = st.columns(2)
                with col1:
                    laba_tahunan = st.number_input("Estimasi Laba Bersih Konsolidasi (Per Tahun)", min_value=0.0, value=500000000.0, step=10000000.0)
                    
                    # Indikator Real-time Laba Tahunan (Format Titik)
                    laba_tahunan_str = f"Rp {laba_tahunan:,.0f}".replace(",", ".")
                    st.success(f"Nominal Terbaca: **{laba_tahunan_str}**")
                    
                    laba_bulanan = laba_tahunan / 12
                    laba_bulanan_str = f"Rp {laba_bulanan:,.0f}".replace(",", ".")
                    st.info(f"Laba Bersih Per Bulan: **{laba_bulanan_str}**")
                
                with col2:
                    limit_baru = st.number_input("Plafon Kredit Baru", min_value=0.0, value=1000000000.0, step=50000000.0)
                    
                    # Indikator Real-time Plafon Baru (Format Titik)
                    limit_baru_str = f"Rp {limit_baru:,.0f}".replace(",", ".")
                    st.success(f"Nominal Terbaca: **{limit_baru_str}**")
                    
                    tenor_baru = st.number_input("Tenor Kredit Baru (Bulan)", min_value=1, value=60, step=12)
                    bunga_baru = st.number_input("Suku Bunga Kredit Baru (% p.a)", min_value=0.0, value=12.0, step=0.5)
                    
                # --- TOMBOL HITUNG ---
                if st.button("Hitung Simulasi DSCR", type="primary"):
                    st.session_state['sudah_dihitung'] = True # Menandai bahwa proses sudah jalan
                    
                    rate_baru_bln = (bunga_baru / 100) / 12
                    if rate_baru_bln > 0:
                        angsuran_baru_anuitas = (limit_baru * rate_baru_bln * (1 + rate_baru_bln)**tenor_baru) / ((1 + rate_baru_bln)**tenor_baru - 1)
                    else:
                        angsuran_baru_anuitas = limit_baru / tenor_baru
                        
                    pokok_rata_bln = limit_baru / tenor_baru
                    current_date = datetime.date(2026, 9, 1) 
                    
                    calc_rows = []
                    for idx, row in edited_df.iterrows():
                        if row['Dihitung?']:
                            limit_asli = float(row['Limit_Str'].replace('Rp', '').replace(',', '').strip())
                            baki_debet_asli = float(row['BakiDebet_Str'].replace('Rp', '').replace(',', '').strip())
                            
                            calc_rows.append({
                                'Kategori': 'Fasilitas Eksisting',
                                'Nama Debitur': row['Nama Debitur'],
                                'Bank - Fasilitas': f"{row['Bank']} - {row['Fasilitas']}",
                                'Jenis': row['Jenis (Rev/Non-Rev)'],
                                'Limit Asli': limit_asli,
                                'Suku Bunga Asli': row['Suku Bunga'],
                                'Awal Kredit': row['Awal Kredit'],
                                'Jatuh Tempo': row['Jatuh Tempo'],
                                'Bulan_Pokok': {},
                                'Bulan_Bunga': {},
                                'Bulan_Total': {},
                                'current_balance': baki_debet_asli
                            })
                    
                    proyeksi_data = []
                    rekap_tahunan_dict = {}
                    bulan_kolom = []
                    
                    balance_baru_anu = limit_baru
                    balance_baru_pr = limit_baru
                    
                    for i in range(int(tenor_baru)):
                        proj_date = current_date + relativedelta(months=i)
                        bulan_label = proj_date.strftime('%b %Y')
                        tahun_label = proj_date.strftime('%Y')
                        bulan_kolom.append(bulan_label)
                        
                        if tahun_label not in rekap_tahunan_dict:
                            rekap_tahunan_dict[tahun_label] = {
                                'Tahun': tahun_label,
                                'Total Pokok (Eksisting + Baru Anuitas)': 0,
                                'Total Bunga (Eksisting + Baru Anuitas)': 0,
                                'Total Angsuran (Eksisting + Baru Anuitas)': 0
                            }
                        
                        total_pokok_eksisting = 0
                        total_bunga_eksisting = 0
                        
                        for crow in calc_rows:
                            try:
                                jt_date = datetime.datetime.strptime(str(crow['Jatuh Tempo']).strip(), '%d/%m/%Y').date()
                                ak_date = datetime.datetime.strptime(str(crow['Awal Kredit']).strip(), '%d/%m/%Y').date()
                            except:
                                jt_date = datetime.date(2099, 12, 31)
                                ak_date = datetime.date(2000, 1, 1)
                                
                            limit_fac = crow['Limit Asli']
                            bunga_fac = float(crow['Suku Bunga Asli']) if pd.notna(crow['Suku Bunga Asli']) else 0
                            rate_bln = (bunga_fac / 100) / 12
                            
                            pokok, bunga, total = 0, 0, 0
                            
                            if proj_date <= jt_date and crow['current_balance'] > 0:
                                if crow['Jenis'] == 'Rev':
                                    bunga = limit_fac * rate_bln
                                    pokok = 0
                                else:
                                    selisih_hari = (jt_date - ak_date).days
                                    tenor_fac = round(selisih_hari / 30.44)
                                    if tenor_fac > 0 and rate_bln > 0:
                                        pmt_fac = (limit_fac * rate_bln * (1 + rate_bln)**tenor_fac) / ((1 + rate_bln)**tenor_fac - 1)
                                    else:
                                        pmt_fac = limit_fac / tenor_fac if tenor_fac > 0 else 0
                                        
                                    bunga = crow['current_balance'] * rate_bln
                                    pokok = max(0, pmt_fac - bunga)
                                    if pokok > crow['current_balance']: pokok = crow['current_balance']
                                    crow['current_balance'] -= pokok
                                
                                total = pokok + bunga
                                
                            crow['Bulan_Pokok'][bulan_label] = pokok
                            crow['Bulan_Bunga'][bulan_label] = bunga
                            crow['Bulan_Total'][bulan_label] = total
                            total_pokok_eksisting += pokok
                            total_bunga_eksisting += bunga
                            
                        bunga_baru_anu_bln = balance_baru_anu * rate_baru_bln
                        pokok_baru_anu_bln = max(0, angsuran_baru_anuitas - bunga_baru_anu_bln)
                        if pokok_baru_anu_bln > balance_baru_anu: pokok_baru_anu_bln = balance_baru_anu
                        total_baru_anu_bln = pokok_baru_anu_bln + bunga_baru_anu_bln
                        balance_baru_anu -= pokok_baru_anu_bln
                        
                        bunga_baru_pr_bln = balance_baru_pr * rate_baru_bln
                        pokok_baru_pr_bln = pokok_rata_bln
                        if pokok_baru_pr_bln > balance_baru_pr: pokok_baru_pr_bln = balance_baru_pr
                        total_baru_pr_bln = pokok_baru_pr_bln + bunga_baru_pr_bln
                        balance_baru_pr -= pokok_baru_pr_bln
                        
                        total_angsuran_eksisting = total_pokok_eksisting + total_bunga_eksisting
                        total_kewajiban_anu = total_angsuran_eksisting + total_baru_anu_bln
                        total_kewajiban_pr = total_angsuran_eksisting + total_baru_pr_bln
                        
                        dscr_anu = laba_bulanan / total_kewajiban_anu if total_kewajiban_anu > 0 else 999
                        dscr_pr = laba_bulanan / total_kewajiban_pr if total_kewajiban_pr > 0 else 999
                        
                        rekap_tahunan_dict[tahun_label]['Total Pokok (Eksisting + Baru Anuitas)'] += (total_pokok_eksisting + pokok_baru_anu_bln)
                        rekap_tahunan_dict[tahun_label]['Total Bunga (Eksisting + Baru Anuitas)'] += (total_bunga_eksisting + bunga_baru_anu_bln)
                        rekap_tahunan_dict[tahun_label]['Total Angsuran (Eksisting + Baru Anuitas)'] += total_kewajiban_anu
                        
                        proyeksi_data.append({
                            'Bulan Ke': i + 1, 'Tahun-Bulan': bulan_label,
                            'Total Pokok Eksisting': total_pokok_eksisting, 'Total Bunga Eksisting': total_bunga_eksisting, 'Total Angsuran Eksisting': total_angsuran_eksisting,
                            'Baru (Anuitas) - Pokok': pokok_baru_anu_bln, 'Baru (Anuitas) - Bunga': bunga_baru_anu_bln, 'Baru (Anuitas) - Total': total_baru_anu_bln,
                            'Baru (Pokok Rata) - Pokok': pokok_baru_pr_bln, 'Baru (Pokok Rata) - Bunga': bunga_baru_pr_bln, 'Baru (Pokok Rata) - Total': total_baru_pr_bln,
                            'DSCR Anuitas': dscr_anu, 'DSCR Pokok Rata': dscr_pr
                        })
                    
                    df_proyeksi = pd.DataFrame(proyeksi_data)
                    df_rekap_tahunan = pd.DataFrame(list(rekap_tahunan_dict.values()))
                    
                    # Simpan data visual ke session_state agar bisa dirender kapan saja
                    st.session_state['df_proyeksi'] = df_proyeksi
                    st.session_state['df_rekap_tahunan'] = df_rekap_tahunan
                    
                    calc_excel_data = []
                    for crow in calc_rows:
                        base_info = {'Nama Debitur': crow['Nama Debitur'], 'Bank - Fasilitas': crow['Bank - Fasilitas'], 'Limit': crow['Limit Asli'], 'Bunga (%)': crow['Suku Bunga Asli'], 'Jatuh Tempo': crow['Jatuh Tempo']}
                        row_pokok = {**base_info, 'Pos / Komponen': 'Angs. Pokok'}
                        row_bunga = {**base_info, 'Pos / Komponen': 'Biaya Bunga'}
                        row_total = {**base_info, 'Pos / Komponen': 'Total Angsuran'}
                        
                        for b_label in bulan_kolom:
                            row_pokok[b_label] = crow['Bulan_Pokok'][b_label]
                            row_bunga[b_label] = crow['Bulan_Bunga'][b_label]
                            row_total[b_label] = crow['Bulan_Total'][b_label]
                            
                        calc_excel_data.extend([row_pokok, row_bunga, row_total])
                    
                    df_calc = pd.DataFrame(calc_excel_data)
                    df_recp = df_proyeksi.drop(columns=['Bulan Ke']).set_index('Tahun-Bulan').T.reset_index().rename(columns={'index': 'Pos / Komponen'})
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_calc.to_excel(writer, sheet_name='CALC', index=False)
                        df_recp.to_excel(writer, sheet_name='RECP', index=False)
                        df_rekap_tahunan.to_excel(writer, sheet_name='REKAP TAHUNAN', index=False)
                        
                        workbook = writer.book
                        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center'})
                        fmt_pokok = workbook.add_format({'num_format': '#,##0', 'bg_color': '#E8F2F8'}) 
                        fmt_bunga = workbook.add_format({'num_format': '#,##0', 'bg_color': '#FCE4D6'}) 
                        fmt_total = workbook.add_format({'num_format': '#,##0', 'bg_color': '#E2EFDA', 'bold': True}) 
                        fmt_default_num = workbook.add_format({'num_format': '#,##0'})
                        
                        for sheet_name, df_sheet in zip(['CALC', 'RECP', 'REKAP TAHUNAN'], [df_calc, df_recp, df_rekap_tahunan]):
                            worksheet = writer.sheets[sheet_name]
                            for col_num, value in enumerate(df_sheet.columns):
                                worksheet.write(0, col_num, value, fmt_header)
                            for i, col in enumerate(df_sheet.columns):
                                max_len = max(df_sheet[col].astype(str).map(len).max(), len(str(col))) + 2
                                worksheet.set_column(i, i, max_len)
                            if 'Pos / Komponen' in df_sheet.columns:
                                pos_idx = list(df_sheet.columns).index('Pos / Komponen')
                                for row_num in range(len(df_sheet)):
                                    val = str(df_sheet.iloc[row_num, pos_idx])
                                    if 'Pokok' in val: fmt_row = fmt_pokok
                                    elif 'Bunga' in val: fmt_row = fmt_bunga
                                    elif 'Total' in val or 'DSCR' in val: fmt_row = fmt_total
                                    else: fmt_row = fmt_default_num
                                    worksheet.set_row(row_num + 1, None, fmt_row)
                            else:
                                for row_num in range(len(df_sheet)):
                                    worksheet.set_row(row_num + 1, None, fmt_default_num)

                    # Simpan file excel mentah ke memori
                    st.session_state['hasil_excel'] = output.getvalue()

                # --- RENDER HASIL & TOMBOL DOWNLOAD DI LUAR BLOK TOMBOL HITUNG ---
                if st.session_state.get('sudah_dihitung'):
                    st.subheader("4. Hasil Simulasi & Proyeksi")
                    df_proyeksi_render = st.session_state['df_proyeksi']
                    
                    col_a, col_b, col_c = st.columns(3)
                    col_a.metric("Total Angsuran Eksisting (Bulan 1)", f"Rp {df_proyeksi_render.iloc[0]['Total Angsuran Eksisting']:,.0f}")
                    col_b.metric("DSCR Kredit Anuitas", f"{df_proyeksi_render.iloc[0]['DSCR Anuitas']:.2f}x")
                    col_c.metric("DSCR Pokok Rata", f"{df_proyeksi_render.iloc[0]['DSCR Pokok Rata']:.2f}x")
                    
                    st.markdown("### Rekapitulasi Tahunan (Skenario Kredit Baru Anuitas)")
                    st.dataframe(st.session_state['df_rekap_tahunan'].style.format({
                        'Total Pokok (Eksisting + Baru Anuitas)': 'Rp {:,.0f}',
                        'Total Bunga (Eksisting + Baru Anuitas)': 'Rp {:,.0f}',
                        'Total Angsuran (Eksisting + Baru Anuitas)': 'Rp {:,.0f}'
                    }), use_container_width=True, hide_index=True)

                    st.write("---") # Garis pembatas
                    st.download_button(
                        label="📥 Download Laporan (CALC & RECP) Excel Rapi",
                        data=st.session_state['hasil_excel'],
                        file_name=f"PampamDSCR_Simulasi_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file atau data: {e}")