import streamlit as st
import pandas as pd
import os

# Set exact wide layout matching the 1200x750 desktop panel size
st.set_page_config(page_title="Laserax Inventory Manager", layout="wide")
st.title("🏭 Laserax GmbH Inventory Manager")

# --- CONFIGURATION: Link to your shared OneDrive Sheet ---
# Paste the direct download link to the inventory.xlsx hosted in your shared OneDrive here
ONEDRIVE_FILE_URL = "https://laseraxinc-my.sharepoint.com/:x:/g/personal/bbhattarai_laserax_com/IQDlpeWcuGCsTKPotwyxsN8fAZVN6H-adOr3sQTjHeCWd5w?e=gpybkY&download=1"

# Columns mapped exactly to your Excel image headers
COLUMNS = ["S.No", "EQUIPMENT", "LASERAX PROJECT No. - Part NO", "STOCK", "LOCATION", "REMARKS", "PROCUREMENT LINK"]

# Cache helper to simulate database rows
if "inventory_df" not in st.session_state:
    try:
        # Pull live sheet from your OneDrive share link
        st.session_state.inventory_df = pd.read_excel(ONEDRIVE_FILE_URL)
        st.session_state.inventory_df.columns = st.session_state.inventory_df.columns.str.strip()
    except:
        # Fallback empty dataframe matching structure if OneDrive link isn't initialized yet
        st.session_state.inventory_df = pd.DataFrame(columns=COLUMNS)

if "S.No" not in st.session_state.inventory_df.columns:
    st.session_state.inventory_df.insert(0, "S.No", range(1, len(st.session_state.inventory_df) + 1))

df = st.session_state.inventory_df

# ==============================================================================
# --- 1. TOP CONTROL FRAME (Search & Reset & Save) ---
# ==============================================================================
st.markdown("---")
top_col1, top_col2, top_col3 = st.columns([3, 1, 2])

with top_col1:
    search_query = st.text_input("Search Equipment:", placeholder="Type name here...", label_visibility="collapsed")
with top_col2:
    search_clicked = st.button("🔍 Search", type="primary", use_container_width=True)
with top_col3:
    reset_clicked = st.button("🔄 Reset Table View", use_container_width=True)

# Apply Search Filters
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
    add_loc = st.text_input("Add Warehouse Location Location:", key="add_loc")
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
        st.success(f"Successfully added '{add_eq}' to row view tracking data structures!")
        st.rerun()
    else:
        st.error("Equipment Name field cannot be left empty.")

# ==============================================================================
# --- 4. MANAGE / EDIT SELECTED ITEM FRAME ---
# ==============================================================================
st.markdown("---")
st.markdown("### ✏️ Modify / Edit Existing Selected Record Data")

if len(df) > 0:
    # Select target row by indexing dropdown selection lists
    select_options = [f"Row {row['S.No']}: {row['EQUIPMENT']}" for _, row in df.iterrows()]
    selected_option = st.selectbox("Choose tracking row record to modify or delete:", select_options)
    
    # Locate exact row matrix items
    selected_sno = int(selected_option.split(":")[0].replace("Row ", ""))
    row_idx = df[df["S.No"] == selected_sno].index[0]
    current_row = df.loc[row_idx]
    
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
            st.session_state.inventory_df.at[row_idx, "STOCK"] += 1
            st.success("Stock increased!")
            st.rerun()
            
    with action_col2:
        if st.button("➖ Decrease Stock (-1)", use_container_width=True):
            current_qty = st.session_state.inventory_df.at[row_idx, "STOCK"]
            st.session_state.inventory_df.at[row_idx, "STOCK"] = max(0, current_qty - 1)
            st.success("Stock decreased!")
            st.rerun()
            
    with action_col3:
        if st.button("💾 Save All Edits", use_container_width=True):
            if edit_eq:
                st.session_state.inventory_df.at[row_idx, "EQUIPMENT"] = edit_eq
                st.session_state.inventory_df.at[row_idx, "LASERAX PROJECT No. - Part NO"] = edit_proj
                st.session_state.inventory_df.at[row_idx, "STOCK"] = edit_stock
                st.session_state.inventory_df.at[row_idx, "LOCATION"] = edit_loc
                st.session_state.inventory_df.at[row_idx, "REMARKS"] = edit_rem
                st.session_state.inventory_df.at[row_idx, "PROCUREMENT LINK"] = edit_link
                st.success("Changes updated locally!")
                st.rerun()
            else:
                st.error("Name field required.")
                
    with action_col4:
        if st.button("🚨 Remove/Delete Row", use_container_width=True):
            st.session_state.inventory_df = df.drop(row_idx).reset_index(drop=True)
            st.session_state.inventory_df["S.No"] = range(1, len(st.session_state.inventory_df) + 1)
            st.warning("Selected tracking line successfully dropped from current views database records!")
            st.rerun()
else:
    st.info("Database table matrix is currently completely blank.")
