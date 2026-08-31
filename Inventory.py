import streamlit as st
import pandas as pd
import openpyxl
import io
import base64
import requests

# Set layout boundaries matching your original desktop manager tool styling
st.set_page_config(page_title="Laserax Inventory Manager", layout="wide")
st.title("🏭 Laserax Inventory Manager")

# ==============================================================================
# --- CONFIGURATION: GitHub Repository Details ---
# ==============================================================================
# Ensure these are accurate and your file inside GitHub is named exactly 'Inventory.xlsx'
GITHUB_USER = r"/BBhattarai445"
GITHUB_REPO = r"Inventory-Management"
FILE_PATH = r"Inventory.xlsx"  

# Secure connection using your Streamlit secret token
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")
API_URL = f"https://github.com{GITHUB_USER}/{GITHUB_REPO}/contents/{FILE_PATH}"
RAW_URL = f"https://github.com/BBhattarai445/Inventory-Management/blob/main/Inventory.xlsx"

def save_to_github(dataframe):
    """
    Automated Save Engine: Commits the updated Excel spreadsheet directly 
    back to your GitHub repository invisibly in the background.
    """
    if not GITHUB_TOKEN:
        st.error("Missing 'GITHUB_TOKEN' secret in your Streamlit dashboard settings.")
        return

    try:
        # 1. Convert the current layout data to an Excel binary block
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            dataframe.to_excel(writer, index=False)
        buffer.seek(0)
        content_encoded = base64.b64encode(buffer.getvalue()).decode()

        # 2. Get the file's current cloud ID version (SHA hash) to allow overwriting
        headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
        get_res = requests.get(API_URL, headers=headers)
        
        sha = ""
        if get_res.status_code == 200:
            try:
                sha = get_res.json().get("sha", "")
            except:
                pass

        # 3. Push the fresh modifications live
        data = {
            "message": "Automated Inventory Sync update via Laserax Web Portal",
            "content": content_encoded,
            "sha": sha
        }
        put_res = requests.put(API_URL, headers=headers, json=data)
        
        # Validates successful creation or update
        if put_res.status_code == 200 or put_res.status_code == 201:
            st.toast("☁️ Repository inventory synced and locked live on GitHub successfully!", icon="✅")
        else:
            st.error(f"GitHub rejected save execution. Status code: {put_res.status_code}")
    except Exception as e:
        st.error(f"Sync issue: {e}")

# Load live records directly from your repository file block
if "inventory_df" not in st.session_state or st.sidebar.button("🔄 Sync Live GitHub Data"):
    try:
        headers = {}
        if GITHUB_TOKEN:
            headers["Authorization"] = f"token {GITHUB_TOKEN}"
        
        # Try fetching via API first
        res = requests.get(API_URL, headers=headers)
        if res.status_code == 200:
            json_data = res.json()
            if isinstance(json_data, dict) and "content" in json_data:
                # Strips out hidden newlines that corrupt the file format reading stream
                raw_base64_str = json_data["content"].replace("\n", "").replace("\r", "").strip()
                file_data = base64.b64decode(raw_base64_str)
                
                # Explicitly utilize openpyxl to resolve string format mapping blocks
                st.session_state.inventory_df = pd.read_excel(io.BytesIO(file_data), engine="openpyxl")
                st.session_state.inventory_df.columns = st.session_state.inventory_df.columns.str.strip()
            else:
                raise ValueError("API response structural error.")
        else:
            # FALLBACK METHOD: Try reading directly via the raw link if API hits a block
            raw_res = requests.get(RAW_URL, headers=headers)
            if raw_res.status_code == 200:
                st.session_state.inventory_df = pd.read_excel(io.BytesIO(raw_res.content), engine="openpyxl")
                st.session_state.inventory_df.columns = st.session_state.inventory_df.columns.str.strip()
            else:
                st.error(f"Could not find file on GitHub. API status: {res.status_code}, Raw URL status: {raw_res.status_code}")
                st.session_state.inventory_df = pd.DataFrame(columns=["S.No", "EQUIPMENT", "LASERAX PROJECT No. - Part NO", "STOCK", "LOCATION", "REMARKS", "PROCUREMENT LINK"])
    except Exception as e:
        st.error(f"Could not parse repository file. Details: {e}")
        st.session_state.inventory_df = pd.DataFrame(columns=["S.No", "EQUIPMENT", "LASERAX PROJECT No. - Part NO", "STOCK", "LOCATION", "REMARKS", "PROCUREMENT LINK"])

if "S.No" not in st.session_state.inventory_df.columns:
    st.session_state.inventory_df.insert(0, "S.No", range(1, len(st.session_state.inventory_df) + 1))

df = st.session_state.inventory_df

# ==============================================================================
# --- 1. TOP CONTROL FRAME (Search & Reset) ---
# ==============================================================================
st.markdown("---")
# FIXED LINE: Added 3 inside st.columns() to prevent column split errors
top_col1, top_col2, top_col3 = st.columns(3)

with top_col1:
    search_query = st.text_input("Search Equipment:", placeholder="Type equipment name to filter dashboard rows...", label_visibility="collapsed")
with top_col2:
    search_clicked = st.button("🔍 Search Table", type="primary", use_container_width=True)
with top_col3:
    reset_clicked = st.button("🔄 Reset View", use_container_width=True)

if search_query and search_clicked:
    display_df = df[df["EQUIPMENT"].astype(str).str.lower().str.contains(search_query.lower())]
else:
    display_df = df

# ==============================================================================
# --- 2. MIDDLE TABLE FRAME (Treeview Data Grid View) ---
# ==============================================================================
st.subheader("📋 Current Stock Inventory Grid View")
st.dataframe(display_df, use_container_width=True, hide_index=True)

# ==============================================================================
# --- 3. ADD NEW ITEM FRAME ---
# ==============================================================================
st.markdown("---")
st.markdown("### ➕ Add New Inventory Line Item")
add_row1_col1, add_row1_col2, add_row1_col3 = st.columns(3)
add_row2_col1, add_row2_col2, add_row2_col3 = st.columns(3)

with add_row1_col1:
    add_eq = st.text_input("Add Equipment Name:", key="add_eq")
with add_row1_col2:
    add_proj = st.text_input("Add Project/Part No:", key="add_proj")
with add_row1_col3:
    add_stock = st.number_input("Add Stock Count Level:", min_value=0, step=1, value=0, key="add_stock")

with add_row2_col1:
    add_loc = st.text_input("Add Warehouse Location:", key="add_loc")
with add_row2_col2:
    add_rem = st.text_input("Add Item Remarks/Notes:", key="add_rem")
with add_row2_col3:
    add_link = st.text_input("Add Procurement URL Link:", key="add_link")

if st.button("Commit Add New Item Line", type="primary"):
    if add_eq:
        next_sno = len(df) + 1
        new_row = pd.DataFrame([{
            "S.No": next_sno, "EQUIPMENT": add_eq, "LASERAX PROJECT No. - Part NO": add_proj,
            "STOCK": add_stock, "LOCATION": add_loc, "REMARKS": add_rem, "PROCUREMENT LINK": add_link
        }])
        st.session_state.inventory_df = pd.concat([df, new_row], ignore_index=True)
        save_to_github(st.session_state.inventory_df)
        st.rerun()
    else:
        st.error("Equipment Name field cannot be left empty.")

# ==============================================================================
# --- 4. MANAGE / EDIT SELECTED ITEM FRAME ---
# ==============================================================================
st.markdown("---")
st.markdown("### ✏️ Modify / Edit Existing Selected Record Data")

if len(df) > 0:
    select_options = [f"Row {row['S.No']}: {row['EQUIPMENT']}" for _, row in df.iterrows()]
    selected_option = st.selectbox("Choose tracking row record to modify or delete:", select_options)
    
    selected_sno = int(selected_option.split(": ").replace("Row ", ""))
    row_idx = df[df["S.No"] == selected_sno].index
    current_row = df.loc[row_idx].iloc[0]
    
    edit_row1_col1, edit_row1_col2, edit_row1_col3 = st.columns(3)
    edit_row2_col1, edit_row2_col2, edit_row2_col3 = st.columns(3)
    
    with edit_row1_col1:
        edit_eq = st.text_input("Equipment Name:", value=str(current_row["EQUIPMENT"]))
    with edit_row1_col2:
        edit_proj = st.text_input("Project No - Part NO:", value=str(current_row["LASERAX PROJECT No. - Part NO"]))
    with edit_row1_col3:
        edit_stock = st.number_input("Current Stock Unit Value:", min_value=0, step=1, value=int(current_row["STOCK"]))
        
    with edit_row2_col1:
        edit_loc = st.text_input("Storage Location Field:", value=str(current_row["LOCATION"]))
    with edit_row2_col2:
        edit_rem = st.text_input("Remarks Logs:", value=str(current_row["REMARKS"]))
    with edit_row2_col3:
        edit_link = st.text_input("Procurement System URL:", value=str(current_row["PROCUREMENT LINK"]))
        
    action_col1, action_col2, action_col3, action_col4 = st.columns(4)
    
    with action_col1:
        if st.button("➕ Increase Stock (+1)", use_container_width=True):
            st.session_state.inventory_df.at[row_idx[0], "STOCK"] += 1
            save_to_github(st.session_state.inventory_df)
            st.rerun()
            
    with action_col2:
        if st.button("➖ Decrease Stock (-1)", use_container_width=True):
            current_qty = st.session_state.inventory_df.at[row_idx[0], "STOCK"]
