```python
import streamlit as st
import pandas as pd
import io
import base64
import requests


# ==============================================================================
# PAGE CONFIGURATION
# ==============================================================================

st.set_page_config(
    page_title="Laserax Inventory Manager",
    layout="wide"
)

st.title("🏭 Laserax Inventory Manager")


# ==============================================================================
# GITHUB CONFIGURATION
# ==============================================================================

GITHUB_USER = "BBhattarai445"
GITHUB_REPO = "Inventory-Management"
FILE_PATH = "Inventory.xlsx"

# GitHub REST API URL
API_URL = (
    f"https://api.github.com/repos/"
    f"{GITHUB_USER}/{GITHUB_REPO}/contents/{FILE_PATH}"
)

# Token stored securely in Streamlit Secrets
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")

# Empty/default dataframe columns
DEFAULT_COLUMNS = [
    "S.No",
    "EQUIPMENT",
    "LASERAX PROJECT No. - Part NO",
    "STOCK",
    "LOCATION",
    "REMARKS",
    "PROCUREMENT LINK"
]


# ==============================================================================
# GITHUB HELPER FUNCTIONS
# ==============================================================================

def get_github_headers():
    """Create headers for GitHub API requests."""

    headers = {
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28"
    }

    if GITHUB_TOKEN:
        headers["Authorization"] = f"Bearer {GITHUB_TOKEN}"

    return headers


def load_from_github():
    """Load Inventory.xlsx from GitHub."""

    if not GITHUB_TOKEN:
        st.error(
            "❌ Missing GITHUB_TOKEN. "
            "Please add GITHUB_TOKEN to your Streamlit Secrets."
        )
        return pd.DataFrame(columns=DEFAULT_COLUMNS)

    try:
        headers = get_github_headers()

        response = requests.get(
            API_URL,
            headers=headers,
            timeout=30
        )

        if response.status_code != 200:
            st.error(
                f"❌ Could not load Inventory.xlsx from GitHub.\n\n"
                f"HTTP Status: {response.status_code}\n"
                f"Response: {response.text[:500]}"
            )

            return pd.DataFrame(columns=DEFAULT_COLUMNS)

        github_data = response.json()

        if "content" not in github_data:
            st.error(
                "❌ GitHub API response did not contain file content."
            )
            return pd.DataFrame(columns=DEFAULT_COLUMNS)

        # GitHub returns Base64 content
        encoded_content = (
            github_data["content"]
            .replace("\n", "")
            .replace("\r", "")
            .strip()
        )

        file_data = base64.b64decode(encoded_content)

        # Read Excel
        dataframe = pd.read_excel(
            io.BytesIO(file_data),
            engine="openpyxl"
        )

        # Clean column names
        dataframe.columns = (
            dataframe.columns
            .astype(str)
            .str.strip()
        )

        # Make sure all required columns exist
        for column in DEFAULT_COLUMNS:
            if column not in dataframe.columns:
                dataframe[column] = ""

        # Keep the expected column order
        dataframe = dataframe[DEFAULT_COLUMNS]

        # Ensure S.No exists and is sequential
        dataframe["S.No"] = range(
            1,
            len(dataframe) + 1
        )

        return dataframe

    except Exception as e:
        st.error(
            f"❌ Error loading inventory from GitHub:\n\n{e}"
        )

        return pd.DataFrame(columns=DEFAULT_COLUMNS)


def save_to_github(dataframe):
    """Save the current dataframe to Inventory.xlsx on GitHub."""

    if not GITHUB_TOKEN:
        st.error(
            "❌ Missing GITHUB_TOKEN. "
            "Please add it to Streamlit Secrets."
        )
        return False

    try:
        # ----------------------------------------------------------------------
        # 1. Create Excel file in memory
        # ----------------------------------------------------------------------

        buffer = io.BytesIO()

        with pd.ExcelWriter(
            buffer,
            engine="openpyxl"
        ) as writer:

            dataframe.to_excel(
                writer,
                index=False,
                sheet_name="Inventory"
            )

        buffer.seek(0)

        # ----------------------------------------------------------------------
        # 2. Convert Excel file to Base64
        # ----------------------------------------------------------------------

        content_encoded = base64.b64encode(
            buffer.getvalue()
        ).decode("utf-8")

        # ----------------------------------------------------------------------
        # 3. Get current file SHA
        # ----------------------------------------------------------------------

        headers = get_github_headers()

        get_response = requests.get(
            API_URL,
            headers=headers,
            timeout=30
        )

        if get_response.status_code != 200:
            st.error(
                f"❌ Could not find Inventory.xlsx on GitHub.\n\n"
                f"HTTP Status: {get_response.status_code}\n"
                f"Response: {get_response.text[:500]}"
            )
            return False

        github_file = get_response.json()

        sha = github_file.get("sha")

        if not sha:
            st.error(
                "❌ GitHub did not return a SHA for Inventory.xlsx."
            )
            return False

        # ----------------------------------------------------------------------
        # 4. Upload updated file
        # ----------------------------------------------------------------------

        payload = {
            "message": "Update inventory via Laserax Inventory Manager",
            "content": content_encoded,
            "sha": sha
        }

        put_response = requests.put(
            API_URL,
            headers=headers,
            json=payload,
            timeout=30
        )

        if put_response.status_code in [200, 201]:

            st.toast(
                "☁️ Inventory successfully saved to GitHub!",
                icon="✅"
            )

            return True

        else:

            st.error(
                f"❌ GitHub rejected the save.\n\n"
                f"HTTP Status: {put_response.status_code}\n"
                f"Response: {put_response.text[:500]}"
            )

            return False

    except Exception as e:

        st.error(
            f"❌ GitHub synchronization error:\n\n{e}"
        )

        return False


# ==============================================================================
# LOAD INVENTORY
# ==============================================================================

if (
    "inventory_df" not in st.session_state
    or st.sidebar.button(
        "🔄 Sync Live GitHub Data"
    )
):

    with st.spinner("Loading inventory from GitHub..."):

        st.session_state.inventory_df = load_from_github()


# Make sure dataframe exists
if "inventory_df" not in st.session_state:
    st.session_state.inventory_df = pd.DataFrame(
        columns=DEFAULT_COLUMNS
    )


df = st.session_state.inventory_df


# ==============================================================================
# ENSURE REQUIRED COLUMNS EXIST
# ==============================================================================

for column in DEFAULT_COLUMNS:

    if column not in df.columns:
        df[column] = ""


# Correct column order
df = df[DEFAULT_COLUMNS]


# Ensure S.No is sequential
df["S.No"] = range(
    1,
    len(df) + 1
)

# Ensure stock is numeric
df["STOCK"] = pd.to_numeric(
    df["STOCK"],
    errors="coerce"
).fillna(0).astype(int)

st.session_state.inventory_df = df


# ==============================================================================
# 1. SEARCH / RESET
# ==============================================================================

st.markdown("---")

st.subheader("🔎 Search Inventory")

top_col1, top_col2, top_col3 = st.columns(
    [2, 1, 1]
)

with top_col1:

    search_query = st.text_input(
        "Search Equipment",
        placeholder="Type equipment name...",
        label_visibility="collapsed"
    )

with top_col2:

    search_clicked = st.button(
        "🔍 Search",
        type="primary",
        use_container_width=True
    )

with top_col3:

    reset_clicked = st.button(
        "🔄 Reset View",
        use_container_width=True
    )


# Reset search
if reset_clicked:

    st.session_state.search_query = ""

    st.rerun()


# Store search
if "search_query" not in st.session_state:
    st.session_state.search_query = ""


if search_clicked:
    st.session_state.search_query = search_query


active_search = st.session_state.search_query


# ==============================================================================
# FILTER TABLE
# ==============================================================================

if active_search:

    display_df = df[
        df["EQUIPMENT"]
        .astype(str)
        .str.contains(
            active_search,
            case=False,
            na=False
        )
    ]

else:

    display_df = df


# ==============================================================================
# 2. INVENTORY TABLE
# ==============================================================================

st.markdown("---")

st.subheader("📋 Current Stock Inventory")

st.dataframe(
    display_df,
    use_container_width=True,
    hide_index=True
)

st.caption(
    f"Showing {len(display_df)} of {len(df)} inventory records."
)


# ==============================================================================
# 3. ADD NEW ITEM
# ==============================================================================

st.markdown("---")

st.subheader("➕ Add New Inventory Line Item")

add_col1, add_col2, add_col3 = st.columns(3)

with add_col1:

    add_eq = st.text_input(
        "Equipment Name",
        key="add_eq"
    )

with add_col2:

    add_proj = st.text_input(
        "Project / Part No.",
        key="add_proj"
    )

with add_col3:

    add_stock = st.number_input(
        "Stock Count",
        min_value=0,
        step=1,
        value=0,
        key="add_stock"
    )


add_col4, add_col5, add_col6 = st.columns(3)

with add_col4:

    add_loc = st.text_input(
        "Warehouse Location",
        key="add_loc"
    )

with add_col5:

    add_rem = st.text_input(
        "Remarks / Notes",
        key="add_rem"
    )

with add_col6:

    add_link = st.text_input(
        "Procurement URL",
        key="add_link"
    )


if st.button(
    "💾 Commit New Inventory Item",
    type="primary",
    use_container_width=True
):

    if not add_eq.strip():

        st.error(
            "❌ Equipment Name cannot be empty."
        )

    else:

        # Create new row
        new_row = pd.DataFrame([{

            "S.No": len(df) + 1,

            "EQUIPMENT": add_eq.strip(),

            "LASERAX PROJECT No. - Part NO":
                add_proj.strip(),

            "STOCK": int(add_stock),

            "LOCATION":
                add_loc.strip(),

            "REMARKS":
                add_rem.strip(),

            "PROCUREMENT LINK":
                add_link.strip()
        }])


        updated_df = pd.concat(
            [df, new_row],
            ignore_index=True
        )


        # Re-number S.No
        updated_df["S.No"] = range(
            1,
            len(updated_df) + 1
        )


        # Save
        if save_to_github(updated_df):

            st.session_state.inventory_df = updated_df

            st.success(
                "✅ New inventory item added successfully."
            )

            st.rerun()


# ==============================================================================
# 4. EDIT / MANAGE EXISTING ITEM
# ==============================================================================

st.markdown("---")

st.subheader(
    "✏️ Modify / Manage Existing Inventory Record"
)


if len(df) == 0:

    st.info(
        "There are currently no inventory records to edit."
    )

else:

    # --------------------------------------------------------------------------
    # Create selection list
    # --------------------------------------------------------------------------

    select_options = [
        f"Row {row['S.No']}: {row['EQUIPMENT']}"
        for _, row in df.iterrows()
    ]


    selected_option = st.selectbox(
        "Choose inventory record",
        select_options
    )


    # --------------------------------------------------------------------------
    # Get selected row
    # --------------------------------------------------------------------------

    try:

        selected_sno = int(
            selected_option
            .split(": ")[0]
            .replace("Row ", "")
            .strip()
        )

        matching_rows = df[
            df["S.No"] == selected_sno
        ]

        if len(matching_rows) == 0:

            st.error(
                "❌ Selected inventory record could not be found."
            )

            row_idx = None

        else:

            row_idx = matching_rows.index[0]

            current_row = df.loc[row_idx]

    except Exception as e:

        st.error(
            f"❌ Selection indexing error: {e}"
        )

        row_idx = None


    # --------------------------------------------------------------------------
    # Edit fields
    # --------------------------------------------------------------------------

    if row_idx is not None:

        edit_col1, edit_col2, edit_col3 = st.columns(3)

        with edit_col1:

            edit_eq = st.text_input(
                "Equipment Name",
                value=str(
                    current_row["EQUIPMENT"]
                ),
                key=f"edit_eq_{selected_sno}"
            )

        with edit_col2:

            edit_proj = st.text_input(
                "Project No. / Part No.",
                value=str(
                    current_row[
                        "LASERAX PROJECT No. - Part NO"
                    ]
                ),
                key=f"edit_proj_{selected_sno}"
            )

        with edit_col3:

            edit_stock = st.number_input(
                "Current Stock",
                min_value=0,
                step=1,
                value=int(
                    current_row["STOCK"]
                ),
                key=f"edit_stock_{selected_sno}"
            )


        edit_col4, edit_col5, edit_col6 = st.columns(3)

        with edit_col4:

            edit_loc = st.text_input(
                "Storage Location",
                value=str(
                    current_row["LOCATION"]
                ),
                key=f"edit_loc_{selected_sno}"
            )

        with edit_col5:

            edit_rem = st.text_input(
                "Remarks",
                value=str(
                    current_row["REMARKS"]
                ),
                key=f"edit_rem_{selected_sno}"
            )

        with edit_col6:

            edit_link = st.text_input(
                "Procurement URL",
                value=str(
                    current_row["PROCUREMENT LINK"]
                ),
                key=f"edit_link_{selected_sno}"
            )


        # ----------------------------------------------------------------------
        # Action buttons
        # ----------------------------------------------------------------------

        action_col1, action_col2, action_col3, action_col4 = st.columns(4)


        # ----------------------------------------------------------------------
        # INCREASE STOCK
        # ----------------------------------------------------------------------

        with action_col1:

            if st.button(
                "➕ Increase Stock (+1)",
                use_container_width=True
            ):

                st.session_state.inventory_df.at[
                    row_idx,
                    "STOCK"
                ] = int(
                    st.session_state.inventory_df.at[
                        row_idx,
                        "STOCK"
                    ]
                ) + 1


                if save_to_github(
                    st.session_state.inventory_df
                ):

                    st.success(
                        "✅ Stock increased by 1."
                    )

                    st.rerun()


        # ----------------------------------------------------------------------
        # DECREASE STOCK
        # ----------------------------------------------------------------------

        with action_col2:

            if st.button(
                "➖ Decrease Stock (-1)",
                use_container_width=True
            ):

                current_qty = int(
                    st.session_state.inventory_df.at[
                        row_idx,
                        "STOCK"
                    ]
                )


                if current_qty > 0:

                    st.session_state.inventory_df.at[
                        row_idx,
                        "STOCK"
                    ] = current_qty - 1


                    if save_to_github(
                        st.session_state.inventory_df
                    ):

                        st.success(
                            "✅ Stock decreased by 1."
                        )

                        st.rerun()

                else:

                    st.warning(
                        "⚠️ Stock is already 0."
                    )


        # ----------------------------------------------------------------------
        # SAVE EDITS
        # ----------------------------------------------------------------------

        with action_col3:

            if st.button(
                "💾 Save Changes",
                type="primary",
                use_container_width=True
            ):

                if not edit_eq.strip():

                    st.error(
                        "❌ Equipment Name cannot be empty."
                    )

                else:

                    updated_df = (
                        st.session_state.inventory_df.copy()
                    )


                    updated_df.at[
                        row_idx,
                        "EQUIPMENT"
                    ] = edit_eq.strip()


                    updated_df.at[
                        row_idx,
                        "LASERAX PROJECT No. - Part NO"
                    ] = edit_proj.strip()


                    updated_df.at[
                        row_idx,
                        "STOCK"
                    ] = int(edit_stock)


                    updated_df.at[
                        row_idx,
                        "LOCATION"
                    ] = edit_loc.strip()


                    updated_df.at[
                        row_idx,
                        "REMARKS"
                    ] = edit_rem.strip()


                    updated_df.at[
                        row_idx,
                        "PROCUREMENT LINK"
                    ] = edit_link.strip()


                    if save_to_github(updated_df):

                        st.session_state.inventory_df = (
                            updated_df
                        )

                        st.success(
                            "✅ Inventory record updated successfully."
                        )

                        st.rerun()


        # ----------------------------------------------------------------------
        # DELETE ITEM
        # ----------------------------------------------------------------------

        with action_col4:

            if st.button(
                "🗑️ Delete Item",
                use_container_width=True
            ):

                updated_df = (
                    st.session_state.inventory_df
                    .drop(index=row_idx)
                    .reset_index(drop=True)
                )


                # Re-number S.No
                updated_df["S.No"] = range(
                    1,
                    len(updated_df) + 1
                )


                if save_to_github(updated_df):

                    st.session_state.inventory_df = (
                        updated_df
                    )

                    st.success(
                        "🗑️ Inventory item deleted successfully."
                    )

                    st.rerun()


# ==============================================================================
# FOOTER / STATUS
# ==============================================================================

st.markdown("---")

st.caption(
    "☁️ Inventory data is synchronized with the GitHub repository."
)

st.caption(
    f"Total inventory records: {len(st.session_state.inventory_df)}"
)
```
