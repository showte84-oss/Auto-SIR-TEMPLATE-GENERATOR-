import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
import json
import io
import os

st.set_page_config(page_title="SGS SIR Report Automation", page_icon="📊", layout="wide")

st.title("📊 SGS Inspection Report (SIR) Excel Generator")
st.markdown("""
ဤ Web App သည် **API Key မလိုဘဲ** SGS Report ပုံများမှ အချက်အလက်များကို `(40 FCL)SIR TEMPLATE.xlsx` ထဲသို့ အလိုအလျောက် ဖြည့်သွင်းပေးမည့် Hybrid စနစ် ဖြစ်ပါသည်။
""")

# Step 1: Prompt Section
with st.expander("📌 အဆင့် (၁) - AI တွင် ပုံများဖတ်ရန် Prompt ယူပါ (နှိပ်၍ဖတ်ပါ)", expanded=True):
    st.markdown("""
    ၁။ အခမဲ့ AI တစ်ခုခု (**Gemini Chat** သို့မဟုတ် **ChatGPT**) ကို ဖွင့်ပါ။  
    ၂။ SGS Report ပုံ (၅) ပုံကို တင်ပြီး အောက်ပါ **Prompt** ကို Copy ကူး၍ ပေးပို့လိုက်ပါ။
    """)
    
    ai_prompt = """Read the attached SGS inspection report images (Stuffing Report, Weight List, Tally Form, VGM Field Report, and Container Survey Report) and extract all data into the following exact JSON format. Return ONLY valid JSON inside a ```json ``` codeblock without any extra text.

{
  "job_ref": "2108339",
  "date": "14/10/2026",
  "start_time": "08:30",
  "shipper": "Yangon Agri Food",
  "buyer": "Kalp Impex",
  "commodity": "Black Eye Beans (BEB)",
  "quantity": "48 MTS (1920 Bags)",
  "packing": "New Single pp Bag",
  "poi": "No. 122, HTY.1",
  "survey_location": "No. 122, HTY.1",
  "vessel": "MV As Susanna",
  "destination": "Valencia, Spain",
  "inspector_names": "Kyaw Nyunt, Myint Ko Ko",
  "test_weight": "10 Kgs",
  "test_weight_no": "014",
  "weighing_percent": 0.1,
  "per_bag_gross": 25.05,
  "per_bag_tare": 0.05,
  "weighing_data": [
    {"page": 1, "bags": 192, "gross_weight": 4809.6}
  ],
  "vgm_other_headers": ["Seal", "SECCO", "Paper"],
  "containers": [
    {
      "no": "CAIU 6562141",
      "condition": "Sound",
      "commence": "08:30",
      "complete": "10:00",
      "bags": 960,
      "gross_mt": 24.048,
      "tare_mt": 0.048,
      "seal": "444072",
      "slings": [100, 100, 100, 100, 100, 100, 100, 100, 100, 60],
      "vgm_others": [0.012, 4.0, 10.0],
      "ctn_tare_kg": 2150,
      "hole_dented": "No",
      "floor_type": "Wood",
      "floor_damaged": "No",
      "free_odour": "Yes",
      "free_infestation": "Yes",
      "clean_dry": "Yes",
      "color": "Red",
      "size": "20'",
      "type": "Steel",
      "suitable": "Yes"
    }
  ]
}

Important Rules for Extraction:
- Format Date as DD/MM/YYYY.
- Survey fields (hole_dented, floor_type, color, type, etc.) must be in Title Case (First letter Capitalized, e.g., "No", "Wood", "Yes", "Red", "Blue", "Steel").
- `slings` array can have up to 10 numbers representing each sling column in Tally Form.
- `vgm_other_headers` must match the exact 3 column headers under 'Other Packages (Kgs)' in the VGM image, and `vgm_others` must contain the 3 corresponding weights."""
    
    st.code(ai_prompt, language="text")

st.divider()

# Step 2: Template & JSON Input
st.subheader("📌 အဆင့် (၂) - Template နှင့် JSON Data ထည့်သွင်းပါ")

col1, col2 = st.columns([1, 2])

with col1:
    default_template = "(40 FCL)SIR TEMPLATE.xlsx"
    use_default = os.path.exists(default_template)
    
    if use_default:
        st.success(f"✅ `{default_template}` ဖိုင်ကို Folder ထဲတွင် အသင့်တွေ့ရှိပါသည်။")
    
    uploaded_template = st.file_uploader(
        "Template ဖိုင်တင်ရန် (Folder ထဲတွင် ရှိပြီးသားဆိုလျှင် ထပ်တင်ရန်မလိုပါ)", 
        type=["xlsx"]
    )

with col2:
    json_input = st.text_area(
        "AI မှ ရရှိလာသော JSON Code ကို ဤနေရာတွင် Paste ချပါ -",
        height=280,
        placeholder='{
  "job_ref": "2108339",
  ...
}'
    )

def set_cell_value(ws, cell_coord, value):
    cell = ws[cell_coord]
    if isinstance(cell, openpyxl.cell.cell.MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell_coord in merged_range:
                top_left_cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                top_left_cell.value = value
                return
    else:
        cell.value = value

def generate_excel(template_stream, data):
    wb = openpyxl.load_workbook(template_stream)
    
    job_ref = str(data.get("job_ref", ""))
    date_slash = data.get("date", "")
    start_time = data.get("start_time", "")
    shipper = data.get("shipper", "")
    buyer = data.get("buyer", "")
    commodity = data.get("commodity", "")
    quantity = data.get("quantity", "")
    packing = data.get("packing", "")
    poi = data.get("poi", "")
    survey_loc = data.get("survey_location", poi)
    vessel = data.get("vessel", "")
    destination = data.get("destination", "")
    inspector_names = data.get("inspector_names", "")
    test_weight = data.get("test_weight", "10 Kgs")
    test_weight_no = data.get("test_weight_no", "")
    weighing_percent = data.get("weighing_percent", 0.1)
    per_bag_gross = data.get("per_bag_gross", 0)
    per_bag_tare = data.get("per_bag_tare", 0)
    weighing_data = data.get("weighing_data", [])
    vgm_headers = data.get("vgm_other_headers", ["Seal", "SECCO", "Paper"])
    containers = data.get("containers", [])
    
    # 1. STUFFING REPORT
    ws_stuffing = wb['STUFFING REPORT']
    set_cell_value(ws_stuffing, 'E5', job_ref)
    set_cell_value(ws_stuffing, 'F6', date_slash)
    set_cell_value(ws_stuffing, 'F7', date_slash)
    set_cell_value(ws_stuffing, 'E8', shipper)
    set_cell_value(ws_stuffing, 'E9', buyer)
    set_cell_value(ws_stuffing, 'E10', commodity)
    set_cell_value(ws_stuffing, 'E11', quantity)
    set_cell_value(ws_stuffing, 'E12', packing)
    set_cell_value(ws_stuffing, 'E13', poi)
    set_cell_value(ws_stuffing, 'E14', vessel)
    set_cell_value(ws_stuffing, 'E15', destination)
    set_cell_value(ws_stuffing, 'E16', inspector_names)
    
    start_row = 21
    for i, c in enumerate(containers):
        row = start_row + i
        set_cell_value(ws_stuffing, f'B{row}', c.get("no", ""))
        set_cell_value(ws_stuffing, f'C{row}', c.get("condition", "SOUND"))
        set_cell_value(ws_stuffing, f'D{row}', c.get("commence", ""))
        set_cell_value(ws_stuffing, f'E{row}', c.get("complete", ""))
        set_cell_value(ws_stuffing, f'F{row}', c.get("bags", 0))
        set_cell_value(ws_stuffing, f'G{row}', c.get("gross_mt", 0))
        set_cell_value(ws_stuffing, f'H{row}', c.get("tare_mt", 0))
        set_cell_value(ws_stuffing, f'I{row}', str(c.get("seal", "")))
        
    for r in range(start_row + len(containers), 61):
        for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
            set_cell_value(ws_stuffing, f'{col}{r}', None)
            
    hide_start_row = start_row + max(10, len(containers))
    for r in range(hide_start_row, 61):
        ws_stuffing.row_dimensions[r].hidden = True
        
    set_cell_value(ws_stuffing, 'C68', test_weight)
    set_cell_value(ws_stuffing, 'G68', test_weight_no)
    
    # 2. WEIGHT LIST (SUMMARY)
    ws_weight = wb['WEIGHT LIST (SUMMARY)']
    set_cell_value(ws_weight, 'C4', job_ref)
    set_cell_value(ws_weight, 'C5', shipper)
    set_cell_value(ws_weight, 'C6', commodity)
    set_cell_value(ws_weight, 'C7', poi)
    set_cell_value(ws_weight, 'C8', weighing_percent)
    
    for i, d in enumerate(weighing_data):
        row = 12 + i
        set_cell_value(ws_weight, f'A{row}', d.get("page", i+1))
        set_cell_value(ws_weight, f'B{row}', d.get("bags", 0))
        set_cell_value(ws_weight, f'C{row}', d.get("gross_weight", 0))
        set_cell_value(ws_weight, f'D{row}', "kgs")
        
    for r in range(12 + len(weighing_data), 27):
        for col in ['A', 'B', 'C', 'D']:
            set_cell_value(ws_weight, f'{col}{r}', None)
            
    set_cell_value(ws_weight, 'G13', per_bag_tare)
    
    # 3. TALLY FORM
    ws_tally = wb['TALLY FORM']
    set_cell_value(ws_tally, 'C4', vessel)
    set_cell_value(ws_tally, 'C5', date_slash)
    set_cell_value(ws_tally, 'C6', shipper)
    
    start_row_tally = 11
    for i, c in enumerate(containers):
        row = start_row_tally + i
        set_cell_value(ws_tally, f'B{row}', c.get("no", ""))
        slings = c.get("slings", [])
        for col_idx in range(10): # 10 sling columns (C to L)
            col_letter = get_column_letter(3 + col_idx)
            val = slings[col_idx] if col_idx < len(slings) else None
            set_cell_value(ws_tally, f'{col_letter}{row}', val)
            
    for r in range(start_row_tally + len(containers), 51):
        for col in range(2, 13):
            set_cell_value(ws_tally, f'{get_column_letter(col)}{r}', None)
            
    hide_start_row_tally = start_row_tally + max(10, len(containers))
    for r in range(hide_start_row_tally, 51):
        ws_tally.row_dimensions[r].hidden = True
        
    # 4. VGM REPORT
    ws_vgm = wb['VGM REPORT']
    set_cell_value(ws_vgm, 'B4', shipper)
    set_cell_value(ws_vgm, 'B5', commodity)
    set_cell_value(ws_vgm, 'B6', vessel)
    set_cell_value(ws_vgm, 'I4', poi)
    set_cell_value(ws_vgm, 'J5', date_slash)
    set_cell_value(ws_vgm, 'L5', start_time)
    
    for idx, h_Col in enumerate(['F9', 'G9', 'H9']):
        if idx < len(vgm_headers):
            set_cell_value(ws_vgm, h_Col, vgm_headers[idx])
            
    start_row_vgm = 10
    for i, c in enumerate(containers):
        row = start_row_vgm + i
        set_cell_value(ws_vgm, f'B{row}', c.get("no", ""))
        set_cell_value(ws_vgm, f'C{row}', c.get("bags", 0))
        set_cell_value(ws_vgm, f'D{row}', per_bag_gross)
        vgm_others = c.get("vgm_others", [0, 0, 0])
        set_cell_value(ws_vgm, f'F{row}', vgm_others[0] if len(vgm_others) > 0 else None)
        set_cell_value(ws_vgm, f'G{row}', vgm_others[1] if len(vgm_others) > 1 else None)
        set_cell_value(ws_vgm, f'H{row}', vgm_others[2] if len(vgm_others) > 2 else None)
        set_cell_value(ws_vgm, f'K{row}', c.get("ctn_tare_kg", 0))
        
    for r in range(start_row_vgm + len(containers), 50):
        for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            set_cell_value(ws_vgm, f'{col}{r}', None)
            
    hide_start_row_vgm = start_row_vgm + max(10, len(containers))
    for r in range(hide_start_row_vgm, 50):
        ws_vgm.row_dimensions[r].hidden = True
        
    # 5. CTN SURVEY REPORT
    ws_survey = wb['CTN SURVEY REPORT']
    set_cell_value(ws_survey, 'B4', shipper)
    set_cell_value(ws_survey, 'B5', commodity)
    set_cell_value(ws_survey, 'B6', quantity)
    set_cell_value(ws_survey, 'I4', job_ref)
    set_cell_value(ws_survey, 'I5', survey_loc)
    set_cell_value(ws_survey, 'I6', date_slash)
    
    start_row_survey = 9
    for i, c in enumerate(containers):
        row = start_row_survey + i
        set_cell_value(ws_survey, f'B{row}', c.get("no", ""))
        set_cell_value(ws_survey, f'C{row}', str(c.get("hole_dented", "No")).title())
        set_cell_value(ws_survey, f'D{row}', str(c.get("floor_type", "Wood")).title())
        set_cell_value(ws_survey, f'E{row}', str(c.get("floor_damaged", "No")).title())
        set_cell_value(ws_survey, f'F{row}', str(c.get("free_odour", "Yes")).title())
        set_cell_value(ws_survey, f'G{row}', str(c.get("free_infestation", "Yes")).title())
        set_cell_value(ws_survey, f'H{row}', str(c.get("clean_dry", "Yes")).title())
        set_cell_value(ws_survey, f'I{row}', str(c.get("color", "")).title())
        set_cell_value(ws_survey, f'J{row}', c.get("size", "20'"))
        set_cell_value(ws_survey, f'K{row}', str(c.get("type", "Steel")).title())
        set_cell_value(ws_survey, f'L{row}', str(c.get("suitable", "Yes")).title())
        
    for r in range(start_row_survey + len(containers), 49):
        for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            set_cell_value(ws_survey, f'{col}{r}', None)
            
    hide_start_row_survey = start_row_survey + max(10, len(containers))
    for r in range(hide_start_row_survey, 49):
        ws_survey.row_dimensions[r].hidden = True
        
    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream, f"SIR_JOB_NO_{job_ref}.xlsx"

# Generate Button
if st.button("🚀 Generate Excel File", type="primary"):
    if not json_input.strip():
        st.error("⚠️ ကျေးဇူးပြု၍ JSON Data ကို အရင်ထည့်သွင်းပေးပါ။")
    else:
        try:
            # Clean markdown codeblock if pasted
            cleaned_json = json_input.strip()
            if cleaned_json.startswith("```"):
                cleaned_json = cleaned_json.split("\n", 1)[1]
                if cleaned_json.endswith("```"):
                    cleaned_json = cleaned_json.rsplit("```", 1)[0]
            
            parsed_data = json.loads(cleaned_json)
            
            template_source = uploaded_template if uploaded_template else (default_template if use_default else None)
            
            if not template_source:
                st.error("⚠️ `(40 FCL)SIR TEMPLATE.xlsx` ဖိုင်ကို Upload တင်ပေးပါ။")
            else:
                excel_bytes, filename = generate_excel(template_source, parsed_data)
                st.success(f"🎉 `{filename}` ဖိုင်ကို အောင်မြင်စွာ ဖန်တီးပြီးပါပြီ!")
                st.download_button(
                    label=f"📥 Download {filename}",
                    data=excel_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except json.JSONDecodeError as e:
            st.error(f"❌ JSON Format မှားယွင်းနေပါသည်: {e}")
        except Exception as e:
            st.error(f"❌ အမှားအယွင်းတစ်ခု ဖြစ်ပေါ်နေပါသည်: {e}")
