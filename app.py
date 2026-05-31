import os
import re
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import streamlit as st

# 1. WEBAPP PAGE INTERFACE SETTINGS
st.set_page_config(page_title="Karad Division MMU Engine", page_icon="📬", layout="centered")

st.title("📬 Karad Division — MMU Report Engine")
st.caption("Department of Posts | India Post, Maharashtra Circle")
st.markdown("---")

st.markdown("### 📥 Upload Daily Data Logs")
st.info("💡 Select your 7 daily transit and compliance files and drop them here simultaneously. The Master Office Directory is permanently handled by the system backend.")

uploaded_files = st.file_uploader(
    "Drop Daily CSV Files Here:", 
    accept_multiple_files=True, 
    type=['csv']
)

# Enforced Absolute Unified Backend Name Strategy
MASTER_FILE_NAME = "Master_Directory.csv"

# 2. OPERATIONAL CORE UTILITIES
def extract_report_date(files_dict):
    """Dynamically parses and enforces the report date from single-day file names"""
    for fname in files_dict.keys():
        match = re.search(r'(\d{2}[\.\-]\d{2}[\.\-]\d{4})', fname)
        if match and ("Productivity" in fname or "DSS" in fname):
            return match.group(1).replace('-', '.')
    return "30.05.2026"  # Fallback default authority date

def calc_pct(disposed, received):
    """Bypasses division-by-zero errors gracefully by returning a hyphen text string"""
    if pd.isna(received) or received == 0:
        return "-"
    return round((disposed / received), 4)

# 3. BACKGROUND PROCESSING ENGINE
if uploaded_files:
    files = {f.name: f for f in uploaded_files}
    
    # Display the number of uploaded files to the user for confirmation
    st.success(f"📂 {len(uploaded_files)} files successfully staged in drop-zone.")
    
    # Intentionally halt automatic compilation loop until user triggers action button explicitly
    if st.button("🚀 Generate Consolidated Report", type="primary"):
        
        # Structural verification check for the permanent master file inside the workspace
        if not os.path.exists(MASTER_FILE_NAME):
            st.error(f"🚨 Backend Alignment Error: The permanent directory file '{MASTER_FILE_NAME}' was not found in the root repository path. Please upload it to your GitHub repository.")
        else:
            with st.spinner("Processing volume aggregation matrices and compiling visual worksheets..."):
                try:
                    # Read Master file with comprehensive multi-encoding fallback safeguards
                    try:
                        master_df = pd.read_csv(MASTER_FILE_NAME, encoding='utf-8-sig')
                    except UnicodeDecodeError:
                        master_df = pd.read_csv(MASTER_FILE_NAME, encoding='cp1252')
                    
                    # Normalize column headers completely (lowercase, strip whitespace, remove punctuation characters)
                    orig_cols = list(master_df.columns)
                    normalized_headers = [str(c).lower().replace("-", " ").replace("_", " ").strip() for c in orig_cols]
                    
                    # Dynamic Key Extraction (Self-Healing Lookup Strategy)
                    col_mapping = {}
                    for idx, h in enumerate(normalized_headers):
                        if "sub division" in h: col_mapping['Sub_Division'] = orig_cols[idx]
                        elif "sub office" in h: col_mapping['Sub_Office'] = orig_cols[idx]
                        elif "branch office" in h or ("office" in h and "name" in h and h != "sub office"): col_mapping['Office_Name'] = orig_cols[idx]
                        elif "office id" in h or "office_id" in h: col_mapping['Office_ID'] = orig_cols[idx]
                        elif "type" in h or "code" in h: col_mapping['Office_Type'] = orig_cols[idx]

                    required_targets = ['Sub_Division', 'Sub_Office', 'Office_Name', 'Office_ID', 'Office_Type']
                    missing_targets = [t for t in required_targets if t not in col_mapping]
                    
                    if missing_targets:
                        st.error(f"🚨 Master Alignment Error: Missing target keys: {missing_targets}. Checked headers: {orig_cols}")
                        st.stop()
                    
                    # Filter and reorder dataset using our safe extracted dictionary maps
                    f_df = pd.DataFrame({
                        'Sub_Division': master_df[col_mapping['Sub_Division']],
                        'Sub_Office': master_df[col_mapping['Sub_Office']],
                        'Office_Name': master_df[col_mapping['Office_Name']],
                        'Office_ID': master_df[col_mapping['Office_ID']],
                        'Office_Type': master_df[col_mapping['Office_Type']]
                    })

                    # Data type coercion to protect joins
                    f_df.dropna(subset=['Office_ID'], inplace=True)
                    f_df['Office_ID'] = f_df['Office_ID'].astype(str).str.strip().str.replace(".0", "", regex=False)
                    f_df = f_df.drop_duplicates(subset=['Office_ID'])
                    
                    # Safe transit logs parsing engine
                    def read_transit_data(keywords):
                        target = [k for k in files.keys() if all(x in k for x in keywords)]
                        if not target:
                            return pd.DataFrame(columns=['office_id', 'Received', 'D0 Delivered', 'D0 Redirected', 'D0 Returned'])
                        
                        try:
                            df_raw = pd.read_csv(files[target[0]], encoding='utf-8-sig')
                        except UnicodeDecodeError:
                            df_raw = pd.read_csv(files[target[0]], encoding='cp1252')
                            
                        df_raw.columns = [str(c).strip() for c in df_raw.columns]
                        id_col = [c for c in df_raw.columns if 'office' in c.lower() and 'id' in c.lower()]
                        
                        if not id_col:
                            id_col = [df_raw.columns[0]]
                            
                        df_raw['office_id'] = df_raw[id_col[0]].astype(str).str.strip().str.replace(".0", "", regex=False)
                        
                        cols = ['Received', 'D0 Delivered', 'D0 Redirected', 'D0 Returned']
                        avail = [c for c in cols if c in df_raw.columns]
                        return df_raw.groupby('office_id')[avail].sum().reset_index()

                    all_prod = read_transit_data(["All Products"])
                    reg_let  = read_transit_data(["Registered Letter"])
                    spd_let  = read_transit_data(["Speed Letter"])
                    reg_par  = read_transit_data(["Registered Parcel"])
                    spd_par  = read_transit_data(["Speed Parcel"])

                    # Combined columns math
                    docs = pd.merge(reg_let, spd_let, on='office_id', how='outer', suffixes=('_r', '_s')).fillna(0)
                    docs['Rec'] = docs.get('Received_r', 0) + docs.get('Received_s', 0)
                    docs['Disp'] = (docs.get('D0 Delivered_r', 0) + docs.get('D0 Redirected_r', 0) + docs.get('D0 Returned_r', 0) +
                                    docs.get('D0 Delivered_s', 0) + docs.get('D0 Redirected_s', 0) + docs.get('D0 Returned_s', 0))

                    parcels = pd.merge(reg_par, spd_par, on='office_id', how='outer', suffixes=('_r', '_s')).fillna(0)
                    parcels['Rec'] = parcels.get('Received_r', 0) + parcels.get('Received_s', 0)
                    parcels['Disp'] = (parcels.get('D0 Delivered_r', 0) + parcels.get('D0 Redirected_r', 0) + parcels.get('D0 Returned_r', 0) +
                                       parcels.get('D0 Delivered_s', 0) + parcels.get('D0 Redirected_s', 0) + parcels.get('D0 Returned_s', 0))

                    # Single snapshots parsing
                    prod_target = [k for k in files.keys() if "Productivity" in k]
                    if prod_target:
                        try:
                            p_df = pd.read_csv(files[prod_target[0]], encoding='utf-8-sig')
                        except UnicodeDecodeError:
                            p_df = pd.read_csv(files[prod_target[0]], encoding='cp1252')
                            
                        p_df.columns = [c.strip() for c in p_df.columns]
                        p_id = [c for c in p_df.columns if 'office' in c.lower() and 'id' in c.lower()][0]
                        p_df['office-id'] = p_df[p_id].astype(str).str.strip().str.replace(".0", "", regex=False)
                        p_df['Prod_Rec'] = p_df['invoice-count']
                        p_df['Prod_Disp'] = p_df['delivery-count'] + p_df['redirection-count'] + p_df['return-count']
                        p_grouped = p_df.groupby('office-id')[['Prod_Rec', 'Prod_Disp']].sum().reset_index()
                    else:
                        p_grouped = pd.DataFrame(columns=['office-id', 'Prod_Rec', 'Prod_Disp'])

                    def get_dss_log(is_daily):
                        t = [k for k in files.keys() if "DSS" in k and ("to" in k if not is_daily else "to" not in k)]
                        if not t:
                            return pd.DataFrame(columns=['office_id', 'pdm', 'dss'])
                        try:
                            df_d = pd.read_csv(files[t[0]], encoding='utf-8-sig')
                        except UnicodeDecodeError:
                            df_d = pd.read_csv(files[t[0]], encoding='cp1252')
                            
                        df_d.columns = [c.strip() for c in df_d.columns]
                        d_id = [c for c in df_d.columns if 'office' in c.lower() and 'id' in c.lower()][0]
                        df_d['office_id'] = df_d[d_id].astype(str).str.strip().str.replace(".0", "", regex=False)
                        return df_d.groupby('office_id')[['total_pdm_art_count', 'total_dss_art_count']].sum().reset_index().rename(
                            columns={'total_pdm_art_count': 'pdm', 'total_dss_art_count': 'dss'}
                        )

                    dss_c = get_dss_log(is_daily=False)
                    dss_d = get_dss_log(is_daily=True)

                    # Base Compilation DataFrame
                    f_df = f_df.merge(all_prod, left_on='Office_ID', right_on='office_id', how='left').drop(columns=['office_id'])
                    f_df['AP_Disp'] = f_df.get('D0 Delivered', 0) + f_df.get('D0 Redirected', 0) + f_df.get('D0 Returned', 0)
                    f_df.rename(columns={'Received': 'AP_Rec'}, inplace=True)

                    f_df = f_df.merge(docs[['office_id', 'Rec', 'Disp']], left_on='Office_ID', right_on='office_id', how='left').drop(columns=['office_id'])
                    f_df.rename(columns={'Rec': 'Doc_Rec', 'Disp': 'Doc_Disp'}, inplace=True)

                    f_df = f_df.merge(parcels[['office_id', 'Rec', 'Disp']], left_on='Office_ID', right_on='office_id', how='left').drop(columns=['office_id'])
                    f_df.rename(columns={'Rec': 'Par_Rec', 'Disp': 'Par_Disp'}, inplace=True)

                    f_df = f_df.merge(p_grouped, left_on='Office_ID', right_on='office-id', how='left').drop(columns=['office-id'])
                    f_df = f_df.merge(dss_c, left_on='Office_ID', right_on='office_id', how='left').drop(columns=['office_id']).rename(columns={'pdm': 'DSS_C_Pdm', 'dss': 'DSS_C_Dss'})
                    f_df = f_df.merge(dss_d, left_on='Office_ID', right_on='office_id', how='left').drop(columns=['office_id']).rename(columns={'pdm': 'DSS_D_Pdm', 'dss': 'DSS_D_Dss'})
                    f_df.fillna(0, inplace=True)

                    # Run baseline formulas
                    f_df['AP_Pct'] = f_df.apply(lambda r: calc_pct(r['AP_Disp'], r['AP_Rec']), axis=1)
                    f_df['Doc_Pct'] = f_df.apply(lambda r: calc_pct(r['Doc_Disp'], r['Doc_Rec']), axis=1)
                    f_df['Par_Pct'] = f_df.apply(lambda r: calc_pct(r['Par_Disp'], r['Par_Rec']), axis=1)
                    f_df['Prod_Pct'] = f_df.apply(lambda r: calc_pct(r['Prod_Disp'], r['Prod_Rec']), axis=1)
                    f_df['DSS_C_Pct'] = f_df.apply(lambda r: calc_pct(r['DSS_C_Dss'], r['DSS_C_Pdm']), axis=1)
                    f_df['DSS_D_Pct'] = f_df.apply(lambda r: calc_pct(r['DSS_D_Dss'], r['DSS_D_Pdm']), axis=1)

                    # 4. EXCEL GENERATION PASSTHROUGHS
                    wb = openpyxl.Workbook()
                    wb.remove(wb.active)

                    f_family = "Arial"
                    font_title = Font(name=f_family, size=14, bold=True, color='FFFFFF')
                    font_header = Font(name=f_family, size=10, bold=True, color='FFFFFF')
                    font_sub_total = Font(name=f_family, size=9, bold=True, color='1F3864')
                    font_grand_total = Font(name=f_family, size=10, bold=True, color='FFFFFF')
                    font_data_reg = Font(name=f_family, size=9, color='000000')
                    font_data_alert = Font(name=f_family, size=9, bold=True, color='B22222')
                    font_data_good = Font(name=f_family, size=9, bold=True, color='1A7A3F')

                    fill_title = PatternFill(fill_type='solid', start_color='1F3864')
                    fill_super = PatternFill(fill_type='solid', start_color='2E75B6')
                    fill_header = PatternFill(fill_type='solid', start_color='595959')
                    fill_sub_div = PatternFill(fill_type='solid', start_color='7B0000')
                    fill_sub_total = PatternFill(fill_type='solid', start_color='FDE8E8')
                    fill_grand_total = PatternFill(fill_type='solid', start_color='7B0000')
                    fill_odd = PatternFill(fill_type='solid', start_color='FFF3F3')
                    fill_even = PatternFill(fill_type='solid', start_color='FFFFFF')
                    fill_alert_cell = PatternFill(fill_type='solid', start_color='FFCCCC')

                    thin_border = Border(left=Side(style='thin', color='AAAAAA'), right=Side(style='thin', color='AAAAAA'), top=Side(style='thin', color='AAAAAA'), bottom=Side(style='thin', color='AAAAAA'))
                    thick_border = Border(left=Side(style='medium', color='888888'), right=Side(style='medium', color='888888'), top=Side(style='medium', color='888888'), bottom=Side(style='medium', color='888888'))

                    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    align_left = Alignment(horizontal='left', vertical='center')

                    rep_date = extract_report_date(files)

                    def setup_headers(ws, title_text):
                        ws.sheet_view.showGridLines = False
                        ws.append([title_text] + [""] * 14)
                        ws.merge_cells('A1:O1')
                        ws.cell(row=1, column=1).font = font_title
                        ws.cell(row=1, column=1).fill = fill_title
                        ws.cell(row=1, column=1).alignment = align_center
                        ws.row_dimensions[1].height = 32

                        r2 = [""] * 5 + ["Delivery Transit Analysis (Range Data)"] * 6 + [f"Delivery Productivity % ({rep_date})"] * 2 + ["DSS Usage Tracking"] * 2
                        ws.append(r2)
                        ws.merge_cells('F2:K2')
                        ws.merge_cells('L2:M2')
                        ws.merge_cells('N2:O2')
                        ws.row_dimensions[2].height = 24

                        r3 = [""] * 5 + ["All Products", "", "Documents", "", "Parcel", "", "", "", "Cumulative Volume", "Daily snapshot"]
                        ws.append(r3)
                        ws.merge_cells('F3:G3')
                        ws.merge_cells('H3:I3')
                        ws.merge_cells('J3:K3')
                        ws.row_dimensions[3].height = 22

                        for r in [2, 3]:
                            for col_idx in range(1, 16):
                                cell = ws.cell(row=r, column=col_idx)
                                if col_idx >= 6:
                                    cell.fill = fill_super
                                    cell.font = font_header
                                    cell.alignment = align_center

                        r4 = ["Sr No.", "Sub Division Name", "Sub Office Name", "Office Name", "Office Type", 
                              "Received", "D+0 %", "Received", "D+0 %", "Received", "D+0 %", 
                              "Received", "D+0 %", "Cumulative (%)", "Daily (%)"]
                        ws.append(r4)
                        for col_idx in range(1, 16):
                            cell = ws.cell(row=4, column=col_idx)
                            cell.fill = fill_header
                            cell.font = font_header
                            cell.alignment = align_center
                            cell.border = thin_border
                        ws.row_dimensions[4].height = 28

                    def format_cell(cell, val, is_p, is_o, dss_col=False):
                        cell.border = thin_border
                        cell.fill = fill_odd if is_o else fill_even
                        if isinstance(val, (int, float)):
                            cell.alignment = align_center
                            if is_p:
                                cell.number_format = '0.0%'
                                limit = 0.80 if dss_col else 0.90
                                if val < limit:
                                    cell.fill = fill_alert_cell
                                    cell.font = font_data_alert
                                else:
                                    cell.font = font_data_good
                            else:
                                cell.number_format = '#,##0'
                                cell.font = font_data_reg
                        else:
                            cell.alignment = align_center if (is_p or val == "-") else align_left
                            cell.font = font_data_reg

                    def set_widths(ws):
                        w_map = {'A': 8, 'B': 22, 'C': 26, 'D': 28, 'E': 12, 'F': 12, 'G': 14, 'H': 12, 'I': 14, 'J': 12, 'K': 14, 'L': 12, 'M': 14, 'N': 16, 'O': 14}
                        for k, v in w_map.items():
                            ws.column_dimensions[k].width = v

                    # SHEET 1: RAW DATA
                    ws1 = wb.create_sheet(title="Raw Data")
                    setup_headers(ws1, f"Consolidated MMU Report (All Active Offices) — {rep_date}")
                    sorted_df = f_df.sort_values(by=['Sub_Division', 'Sub_Office', 'Office_Name'])
                    
                    curr_row, line_odd = 5, True
                    for sub_name, g in sorted_df.groupby('Sub_Division'):
                        s_no = 1
                        for _, row in g.iterrows():
                            v = [s_no, row['Sub_Division'], row['Sub_Office'], row['Office_Name'], row['Office_Type'],
                                 row['AP_Rec'], row['AP_Pct'], row['Doc_Rec'], row['Doc_Pct'], row['Par_Rec'], row['Par_Pct'],
                                 row['Prod_Rec'], row['Prod_Pct'], row['DSS_C_Pct'], row['DSS_D_Pct']]
                            ws1.append(v)
                            ws1.row_dimensions[curr_row].height = 19
                            for col in range(1, 16):
                                format_cell(ws1.cell(row=curr_row, column=col), v[col-1], (col in [7,9,11,13,14,15]), line_odd, dss_col=(col==14))
                            s_no += 1
                            curr_row += 1
                            line_odd = not line_odd
                    set_widths(ws1)
                    ws1.freeze_panes = "F5"

                    # SHEET 2 & 3: FORMATED GROUP TABS
                    def render_formatted_tab(t_name, title_banner, bo_filter):
                        ws = wb.create_sheet(title=t_name)
                        setup_headers(ws, title_banner)
                        
                        df_sub = sorted_df[sorted_df['Office_Type'] != 'BPO'] if bo_filter == 'EX' else sorted_df[sorted_df['Office_Type'] == 'BPO']
                        r_pos, is_row_odd = 5, True
                        
                        gt = {k: 0 for k in ['ap_r','ap_d','doc_r','doc_d','par_r','par_d','pr_r','pr_d','dc_p','dc_d','dd_p','dd_d']}
                        
                        for s_div, g in df_sub.groupby('Sub_Division'):
                            ws.append(["", f"  ▌ {s_div}"] + [""] * 13)
                            ws.merge_cells(start_row=r_pos, start_column=2, end_row=r_pos, end_column=5)
                            ws.row_dimensions[r_pos].height = 21
                            for col in range(1, 16):
                                cell_obj = ws.cell(row=r_pos, column=col)
                                cell_obj.fill = fill_sub_div
                                cell_obj.font = font_header
                                if col == 2: cell_obj.alignment = align_left
                            r_pos += 1
                            
                            s_no = 1
                            st = {k: 0 for k in ['ap_r','ap_d','doc_r','doc_d','par_r','par_d','pr_r','pr_d','dc_p','dc_d','dd_p','dd_d']}
                            
                            for _, row in g.iterrows():
                                v = [s_no, row['Sub_Division'], row['Sub_Office'], row['Office_Name'], row['Office_Type'],
                                     row['AP_Rec'], row['AP_Pct'], row['Doc_Rec'], row['Doc_Pct'], row['Par_Rec'], row['Par_Pct'],
                                     row['Prod_Rec'], row['Prod_Pct'], row['DSS_C_Pct'], row['DSS_D_Pct']]
                                ws.append(v)
                                ws.row_dimensions[r_pos].height = 19
                                for col in range(1, 16):
                                    format_cell(ws.cell(row=r_pos, column=col), v[col-1], (col in [7,9,11,13,14,15]), is_row_odd, dss_col=(col==14))
                                
                                st['ap_r'] += row['AP_Rec']
                                st['ap_d'] += row['AP_Disp']
                                st['doc_r'] += row['Doc_Rec']
                                st['doc_d'] += row['Doc_Disp']
                                st['par_r'] += row['Par_Rec']
                                st['par_d'] += row['Par_Disp']
                                st['pr_r']  += row['Prod_Rec']
                                st['pr_d']  += row['Prod_Disp']
                                st['dc_p']  += row['DSS_C_Pdm']
                                st['dc_d']  += row['DSS_C_Dss']
                                st['dd_p']  += row['DSS_D_Pdm']
                                st['dd_d']  += row['DSS_D_Dss']
                                
                                s_no += 1
                                r_pos += 1
                                is_row_odd = not is_row_odd

                            # Weighted Subtotal Row Math
                            st_row = ["", s_div, f"{s_div} (Total)", "", "",
                                      st['ap_r'], (st['ap_d']/st['ap_r'] if st['ap_r']>0 else "-"),
                                      st['doc_r'], (st['doc_d']/st['doc_r'] if st['doc_r']>0 else "-"),
                                      st['par_r'], (st['par_d']/st['par_r'] if st['par_r']>0 else "-"),
                                      st['pr_r'], (st['pr_d']/st['pr_r'] if st['pr_r']>0 else "-"),
                                      (st['dc_d']/st['dc_p'] if st['dc_p']>0 else "-"),
                                      (st['dd_d']/st['dd_p'] if st['dd_p']>0 else "-")]
                            ws.append(st_row)
                            ws.merge_cells(start_row=r_pos, start_column=3, end_row=r_pos, end_column=4)
                            ws.row_dimensions[r_pos].height = 21
                            for col in range(1, 16):
                                cell = ws.cell(row=r_pos, column=col)
                                cell.fill = fill_sub_total
                                cell.font = font_sub_total
                                cell.border = thick_border
                                if col >= 6:
                                    cell.alignment = align_center
                                    cell.number_format = '0.0%' if col in [7,9,11,13,14,15] else '#,##0'
                            
                            for k in gt: gt[k] += st[k]
                            r_pos += 1
                            
                            # Visual buffer spacing blank row
                            ws.append([""] * 15)
                            ws.row_dimensions[r_pos].height = 8
                            r_pos += 1

                        # Master Grand Total Row Math
                        gt_row = ["", "Karad Division Total", "", "", "",
                                  gt['ap_r'], (gt['ap_d']/gt['ap_r'] if gt['ap_r']>0 else "-"),
                                  gt['doc_r'], (gt['doc_d']/gt['doc_r'] if gt['doc_r']>0 else "-"),
                                  gt['par_r'], (gt['par_d']/gt['par_r'] if gt['par_r']>0 else "-"),
                                  gt['pr_r'], (gt['pr_d']/gt['pr_r'] if gt['pr_r']>0 else "-"),
                                  (gt['dc_d']/gt['dc_p'] if gt['dc_p']>0 else "-"),
                                  (gt['dd_d']/gt['dd_p'] if gt['dd_p']>0 else "-")]
                        ws.append(gt_row)
                        ws.merge_cells(start_row=r_pos, start_column=2, end_row=r_pos, end_column=5)
                        ws.row_dimensions[r_pos].height = 24
                        for col in range(1, 16):
                            cell = ws.cell(row=r_pos, column=col)
                            cell.fill = fill_grand_total
                            cell.font = font_grand_total
                            cell.border = thick_border
                            if col >= 6:
                                cell.alignment = align_center
                                cell.number_format = '0.0%' if col in [7,9,11,13,14,15] else '#,##0'
                        set_widths(wb.active) # Reference update
                        
                    set_widths(ws)
                    ws.freeze_panes = "F5"
                    
                    # Return overall aggregated rates for the on-screen live Streamlit snapshot display layout
                    div_summary_metrics = {
                        'ap_rec': gt['ap_r'],
                        'ap_pct': (gt['ap_d'] / gt['ap_r'] if gt['ap_r'] > 0 else 0),
                        'prod_rec': gt['pr_r'],
                        'prod_pct': (gt['pr_d'] / gt['pr_r'] if gt['pr_r'] > 0 else 0),
                        'dss_d_pct': (gt['dd_d'] / gt['dd_p'] if gt['dd_p'] > 0 else 0),
                        'dss_c_pct': (gt['dc_d'] / gt['dc_p'] if gt['dc_p'] > 0 else 0)
                    }
                    return div_summary_metrics

                m_ex = render_formatted_tab("SPO & HPO", f"Consolidated MMU Report (Excluding B.Os) — {rep_date}", 'EX')
                render_formatted_tab("Only BOs", f"Consolidated MMU Report (Only Branch Offices) — {rep_date}", 'ONLY')

                # SHEET 4: EXECUTIVE AT A GLANCE MANAGEMENT SUMMARY
                ws4 = wb.create_sheet(title="At A Glance")
                ws4.sheet_view.showGridLines = False
                ws4.row_dimensions[1].height = 32
                ws4.row_dimensions[2].height = 22
                ws4.row_dimensions[3].height = 26

                ws4.append([f"Sub Division Wise Executive Performance Summary — {rep_date}"] + [""] * 13)
                ws4.merge_cells("A1:N1")
                ws4.cell(row=1, column=1).font = font_title
                ws4.cell(row=1, column=1).fill = fill_title
                ws4.cell(row=1, column=1).alignment = align_center

                ws4.append(["", "", "Excluding Branch Offices (HPOs & SPOs)", "", "", "", "", "", "Only Branch Offices (BPOs)", "", "", "", "", ""])
                ws4.merge_cells("C2:H2")
                ws4.merge_cells("I2:N2")
                for col in range(1, 15):
                    cell = ws4.cell(row=2, column=col)
                    if col >= 3:
                        cell.fill = fill_super
                        cell.font = font_header
                        cell.alignment = align_center

                ws4_h = ["Sr No.", "Sub Division Name", "Total Offices", "Docs Rec", "Docs D+0 %", "Par Rec", "Par D+0 %", "DSS Cum %",
                         "Total Offices", "Docs Rec", "Docs D+0 %", "Par Rec", "Par D+0 %", "DSS Cum %"]
                ws4.append(ws4_h)
                for col in range(1, 15):
                    cell = ws4.cell(row=3, column=col)
                    cell.fill = fill_header
                    cell.font = font_header
                    cell.alignment = align_center
                    cell.border = thin_border

                sub_div_list = sorted(list(f_df['Sub_Division'].unique()))
                idx_s, odd_line = 4, True
                for i, s_div in enumerate(sub_div_list, 1):
                    df_ex = f_df[(f_df['Sub_Division'] == s_div) & (f_df['Office_Type'] != 'BPO')]
                    df_bo = f_df[(f_df['Sub_Division'] == s_div) & (f_df['Office_Type'] == 'BPO')]
                    
                    row_v = [
                        i, s_div,
                        len(df_ex), df_ex['Doc_Rec'].sum(), (df_ex['Doc_Disp'].sum()/df_ex['Doc_Rec'].sum() if df_ex['Doc_Rec'].sum()>0 else "-"), df_ex['Par_Rec'].sum(), (df_ex['Par_Disp'].sum()/df_ex['Par_Rec'].sum() if df_ex['Par_Rec'].sum()>0 else "-"), (df_ex['DSS_C_Dss'].sum()/df_ex['DSS_C_Pdm'].sum() if df_ex['DSS_C_Pdm'].sum()>0 else "-"),
                        len(df_bo), df_bo['Doc_Rec'].sum(), (df_bo['Doc_Disp'].sum()/df_bo['Doc_Rec'].sum() if df_bo['Doc_Rec'].sum()>0 else "-"), df_bo['Par_Rec'].sum(), (df_bo['Par_Disp'].sum()/df_bo['Par_Rec'].sum() if df_bo['Par_Rec'].sum()>0 else "-"), (df_bo['DSS_C_Dss'].sum()/df_bo['DSS_C_Pdm'].sum() if df_bo['DSS_C_Pdm'].sum()>0 else "-")
                    ]
                    ws4.append(row_v)
                    ws4.row_dimensions[idx_s].height = 20
                    for col in range(1, 15):
                        format_cell(ws4.cell(row=idx_s, column=col), row_v[col-1], (col in [5,7,8,11,13,14]), odd_line, dss_col=(col in [8,14]))
                    idx_s += 1
                    odd_line = not odd_line

                w_sum = {'A': 8, 'B': 22, 'C': 12, 'D': 12, 'E': 14, 'F': 12, 'G': 14, 'H': 14, 'I': 12, 'J': 12, 'K': 14, 'L': 12, 'M': 14, 'N': 14}
                for k, v in w_sum.items():
                    ws4.column_dimensions[k].width = v

                # SHEET 5: ACTIONS DEFAULTER AUDIT LIST WITH STR/FLOAT COMPLIANCE DEFENSE
                ws5 = wb.create_sheet(title="Defaulters List")
                setup_headers(ws5, f"Operational KPI Defaulters Audit List — {rep_date}")
                
                def filter_defaulters(row):
                    ap_val = row['AP_Pct']
                    dss_val = row['DSS_C_Pct']
                    if isinstance(ap_val, (int, float)) and ap_val < 0.90:
                        return True
                    if isinstance(dss_val, (int, float)) and dss_val < 0.80:
                        return True
                    return False

                defcheck = f_df[f_df.apply(filter_defaulters, axis=1)].sort_values(by=['Sub_Division', 'Sub_Office', 'Office_Name'])

                d_row, d_sr = 5, 1
                for _, row in defcheck.iterrows():
                    v = [d_sr, row['Sub_Division'], row['Sub_Office'], row['Office_Name'], row['Office_Type'],
                         row['AP_Rec'], row['AP_Pct'], row['Doc_Rec'], row['Doc_Pct'], row['Par_Rec'], row['Par_Pct'],
                         row['Prod_Rec'], row['Prod_Pct'], row['DSS_C_Pct'], row['DSS_D_Pct']]
                    ws5.append(v)
                    ws5.row_dimensions[d_row].height = 19
                    for col in range(1, 16):
                        format_cell(ws5.cell(row=d_row, column=col), v[col-1], (col in [7,9,11,13,14,15]), is_o=True, dss_col=(col==14))
                    d_row += 1
                    d_sr += 1
                set_widths(ws5)
                ws5.freeze_panes = "F5"

                # 5. STREAMLIT ON-SCREEN SNAPSHOT EXECUTIVE VIEW 
                # This injects a clean summary layout section right inside the browser window
                st.markdown(f"### 📊 Karad Division Executive Snapshot — {rep_date}")
                
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Total Range Volume", f"{m_ex['ap_rec']:,}")
                col2.metric("All Products D+0", f"{m_ex['ap_pct']:.1%}")
                col3.metric("Daily Productivity", f"{m_ex['prod_pct']:.1%}")
                col4.metric("Daily DSS Usage", f"{m_ex['dss_d_pct']:.1%}")
                
                # Compliance Alert Callouts based on target metrics
                if m_ex['ap_pct'] < 0.90 or m_ex['dss_d_pct'] < 0.80:
                    st.warning(f"⚠️ Notice: Division performance averages have dropped below baseline targets. {len(defcheck)} active offices flagged in the audit exception sheet.")
                else:
                    st.success("💪 Core Division tracking benchmarks are currently fully optimized.")

                # 6. MEMORY EXPORT STRATEGY FOR WEB DOWNLOAD
                out_name = f"Consolidated_MMU_Report_{rep_date}.xlsx"
                wb.save(out_name)
                
                with open(out_name, "rb") as file_bytes:
                    st.download_button(
                        label="📥 Download Consolidated Excel Report",
                        data=file_bytes,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                os.remove(out_name)

            except Exception as error:
                st.error(f"An operational pipeline compiling error occurred: {str(error)}")
