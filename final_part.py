import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
import io
from deep_translator import GoogleTranslator

st.set_page_config(page_title="Converter", page_icon="🚌")

st.markdown("""
<style>
    .stApp { background: linear-gradient(to right, #eef2ff, #f8fafc) !important; }
    

    .main-title { 
        font-size: 48px !important; 
        font-weight: 700 !important; 
        text-align: center; 
        margin-bottom: 20px;
        color: #3b82f6 !important;  
    }
    
    /* விட்ஜெட் லேபிள்களை தடிமனான கருப்பு நிறத்தில் வைத்திருக்கும் */
    [data-testid="stWidgetLabel"] p {
        color: #000000 !important;     
        font-weight: 800 !important;   
        font-size: 18px !important;    
    }
</style>
""", unsafe_allow_html=True)
    
st.markdown('<div class="main-title">🚌 Filllist To Customer Sheet</div>', unsafe_allow_html=True)
st.info("Upload the Filllist and Route List files.....")

if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0

if 'done' not in st.session_state:
    st.session_state.done = False
if 'output_data' not in st.session_state:
    st.session_state.output_data = None

col1, col2 = st.columns(2)

with col1:
    uploaded_filllist = st.file_uploader(
        "1. Upload filllist (.xls, .xlsx)", 
        type=['xls', 'xlsx'], 
        key=f"filllist_{st.session_state.uploader_key}"
    )

with col2:
    uploaded_routelist = st.file_uploader(
        "2. Upload Route List (.xls, .xlsx)", 
        type=['xls', 'xlsx'], 
        key=f"routelist_{st.session_state.uploader_key}"
    )

if uploaded_filllist is None or uploaded_routelist is None:
    st.session_state.done = False
    st.session_state.output_data = None

if uploaded_filllist and uploaded_routelist:
    if not st.session_state.done:
        if st.button("🚀 Start Processing"):
            try:
                raw_route_df = pd.read_excel(uploaded_routelist, header=None)
                
                header_row_idx = 0
                r_no_col_idx = 2    # Default Column C
                r_name_col_idx = 3  # Default Column D
                
                for r_idx in range(min(10, len(raw_route_df))):
                    row_vals = [str(x).strip().lower() for x in raw_route_df.iloc[r_idx].dropna()]
                    if any('route' in v and ('no' in v or 'num' in v) for v in row_vals):
                        header_row_idx = r_idx
                        break
                
                route_df = pd.read_excel(uploaded_routelist, header=header_row_idx)
                route_df.columns = [str(c).strip().lower() for c in route_df.columns]
                
                for c_idx, col_name in enumerate(route_df.columns):
                    if 'route' in col_name and ('no' in col_name or 'num' in col_name):
                        r_no_col_idx = c_idx
                    if 'route' in col_name and ('name' in col_name or 'nam' in col_name):
                        r_name_col_idx = c_idx

                route_mapping = {}
                for _, row in route_df.iterrows():
                    if len(row) > max(r_no_col_idx, r_name_col_idx):
                        raw_no = row.iloc[r_no_col_idx]
                        raw_name = row.iloc[r_name_col_idx]
                        
                        if pd.notna(raw_no) and pd.notna(raw_name):
                            if isinstance(raw_no, float) and raw_no.is_integer():
                                r_no_str = str(int(raw_no)).strip().upper()
                            else:
                                r_no_str = str(raw_no).strip().upper()
                                if r_no_str.endswith('.0'):
                                    r_no_str = r_no_str[:-2]
                                    
                            route_mapping[r_no_str] = str(raw_name).strip()

                all_sheets = pd.read_excel(uploaded_filllist, sheet_name=None, header=None)
                translator = GoogleTranslator(source='en', target='ta')
                output = io.BytesIO()
                wb = Workbook()
                wb.remove(wb.active)  
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                     top=Side(style='thin'), bottom=Side(style='thin'))
                
                sheets_list = list(all_sheets.keys())
                total_sheets = len(sheets_list)
                my_bar = st.progress(0, text="Processing please wait...")

                for idx, sheet_key in enumerate(sheets_list):
                    df = all_sheets[sheet_key]
                    if df.empty: 
                        continue
                    
                    ws = wb.create_sheet(title=str(sheet_key))
                    
                    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
                    ws.sheet_properties.pageSetUpPr.fitToPage = True
                    ws.page_setup.fitToWidth = 1
                    ws.page_setup.fitToHeight = False
                    
                    ws.page_margins.left = 0.25
                    ws.page_margins.right = 0.25
                    ws.page_margins.top = 0.5
                    ws.page_margins.bottom = 0.5
                    try:
                        ws.sheet_view.showGridLines = True
                    except Exception:
                        try:
                            ws.views.sheetView[0].showGridLines = True
                        except Exception:
                            try:
                                ws.views.sheetView.showGridLines = True
                            except Exception:
                                pass
                  

                 

                    
                    current_route_no = str(sheet_key).strip().upper()
                    if current_route_no.endswith('.0'):
                        current_route_no = current_route_no[:-2]
                    
                    found_route_name = "N/A"
                    if current_route_no in route_mapping:
                        found_route_name = route_mapping[current_route_no]

                    bus_list = []
                    if df.shape[1] > 1:
                        bus_list = df.iloc[0, 1:].dropna().astype(str).tolist()
                        bus_list = [b.split('.')[0].strip() for b in bus_list if b.strip() != ""]
                    
                    bus_count = len(bus_list) if bus_list else 1
                    bus_no_str = ", ".join(bus_list) if bus_list else "N/A"

                    locations = df.iloc[1:, 0].dropna().astype(str).tolist()
                    translated_locations = []
                    if locations:
                        try:
                            bulk_text = "\n".join(locations)
                            bulk_translated = translator.translate(bulk_text)
                            translated_locations = bulk_translated.split("\n")
                        except Exception:
                            translated_locations = locations

                    ws.merge_cells('A1:A2')
                    ws['A1'] = f"Route No: {current_route_no}"
                    ws['A1'].font = Font(bold=True)
                    ws['A1'].alignment = Alignment(horizontal='left', vertical='center')
                    ws['A1'].border = thin_border
                    ws['A2'].border = thin_border 
                    
                    ws.merge_cells('B1:C2')
                    ws['B1'] = f"{found_route_name}"  
                    ws['B1'].font = Font(bold=True, size=12, color="000000")
                    ws['B1'].alignment = Alignment(horizontal='center', vertical='center')
                    
                    for r in range(1, 3):
                        for c in range(2, 4): 
                            ws.cell(row=r, column=c).border = thin_border
                    
                    ws.merge_cells('B3:C3')
                    ws['B3'] = f"Bus No: {bus_no_str}"  
                    ws['B3'].font = Font(bold=True)
                    ws['B3'].alignment = Alignment(horizontal='center', vertical='center')
                    
                    for c in range(1, 5): 
                        ws.cell(row=3, column=c).border = thin_border
                    ws['A3'].border = thin_border

                    labels = [("No of Singles", 1), ("10", 2), (f"No of Bus-{bus_count}", 3)]
                    for text, r_idx in labels:
                        cell = ws.cell(row=r_idx, column=4, value=text)
                        cell.border = thin_border
                        cell.alignment = Alignment(horizontal='center', vertical='center')

                    headers = ['Slot', 'Ad Location (Tamil)', 'Ad Location (English)', '']
                    for c_idx, text in enumerate(headers, 1):
                        cell = ws.cell(row=4, column=c_idx, value=text)
                        cell.font = Font(bold=True)
                        cell.border = thin_border
                        cell.alignment = Alignment(horizontal='center')
                    font_bold_main = Font(name='Calibri', size=11, bold=True)
                    font_regular = Font(name='Calibri', size=11, bold=False)

                    for i, eng in enumerate(locations):
                        curr_row = 5 + i
                        tam = translated_locations[i] if i < len(translated_locations) else eng
                        
                        c1 = ws.cell(row=curr_row, column=1, value=f"A_{i+1}")
                        c2 = ws.cell(row=curr_row, column=2, value=tam)
                        c3 = ws.cell(row=curr_row, column=3, value=eng)
                        c4 = ws.cell(row=curr_row, column=4, value="")
                        
                        for cell in [c1, c2, c3, c4]:
                            cell.border = thin_border
                            cell.font = font_regular
                        
                        c1.alignment = Alignment(horizontal='left', vertical='center')
                        c2.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                        c3.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                        c4.alignment = Alignment(horizontal='center', vertical='center')
                        
                        ws.row_dimensions[curr_row].height = 28

                    ws.column_dimensions['A'].width = 18
                    ws.column_dimensions['B'].width = 35
                    ws.column_dimensions['C'].width = 35
                    ws.column_dimensions['D'].width = 15

                    my_bar.progress((idx + 1) / total_sheets)

                wb.save(output)
                st.session_state.output_data = output.getvalue()
                st.session_state.done = True
                st.rerun()

            except Exception as e:
                st.error(f"Error: {e}")
if st.session_state.done:
        st.success("✅ Done!")
        st.download_button(
            label="Download Final Excel",
            data=st.session_state.output_data,
            file_name="Customer_Sheet_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        if st.button("🔄 Process Another File"):
            st.session_state.done = False
            st.session_state.output_data = None
            st.session_state.uploader_key += 1  
            st.rerun()
