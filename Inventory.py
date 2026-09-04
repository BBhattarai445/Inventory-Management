from pathlib import Path
import streamlit as st
import pandas as pd
import io
import base64
import requests



# ==============================================================================
# PAGE CONFIG
# ==============================================================================

LOGO_PATH = Path(__file__).parent / "Picture1.png"

st.set_page_config(
    page_title="Laserax Inventory Manager",
    page_icon=str(LOGO_PATH),
    layout="wide"
)

with open("Picture1.png", "rb") as f:
    logo = base64.b64encode(f.read()).decode()

st.markdown(
    f"""
    <div style="display: flex; justify-content: center; width: 100%;">
        <img src="data:image/png;base64,{logo}" style="width: 500px;">
    </div>
    """,
    unsafe_allow_html=True
)
st.title("🏭 Laserax GmbH Inventory Manager")

# ==============================================================================
# GITHUB CONFIGURATION
# ==============================================================================

GITHUB_USER = "BBhattarai445"
GITHUB_REPO = "Inventory-Management"
FILE_PATH = "Inventory.xlsx"

API_URL = (
    f"https://api.github.com/repos/"
    f"{GITHUB_USER}/{GITHUB_REPO}/contents/{FILE_PATH}"
)


# ==============================================================================
# STREAMLIT SECRETS
# ==============================================================================

GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")


# ==============================================================================
# INVENTORY COLUMNS
# ==============================================================================

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
# SESSION STATE
# ==============================================================================

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "role" not in st.session_state:
    st.session_state.role = None

if "username" not in st.session_state:
    st.session_state.username = None

if "inventory_df" not in st.session_state:
    st.session_state.inventory_df = None


# ==============================================================================
# LOGIN PAGE
# ==============================================================================
# ==============================================================================
# GITHUB HEADERS
# ==============================================================================

def get_github_headers():

    headers = {
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28"
    }

    if GITHUB_TOKEN:

        headers["Authorization"] = (
            f"Bearer {GITHUB_TOKEN}"
        )

    return headers


# ==============================================================================
# LOAD INVENTORY FROM GITHUB
# ==============================================================================

def load_from_github():

    if not GITHUB_TOKEN:

        st.error(
            "❌ GITHUB_TOKEN is missing from Streamlit Secrets."
        )

        return pd.DataFrame(
            columns=DEFAULT_COLUMNS
        )


    try:

        response = requests.get(
            API_URL,
            headers=get_github_headers(),
            timeout=30
        )


        if response.status_code != 200:

            st.error(
                f"❌ Could not load Inventory.xlsx from GitHub.\n\n"
                f"HTTP Status: {response.status_code}\n\n"
                f"{response.text[:500]}"
            )

            return pd.DataFrame(
                columns=DEFAULT_COLUMNS
            )


        github_data = response.json()


        if "content" not in github_data:

            st.error(
                "❌ GitHub response did not contain file content."
            )

            return pd.DataFrame(
                columns=DEFAULT_COLUMNS
            )


        encoded_content = (
            github_data["content"]
            .replace("\n", "")
            .replace("\r", "")
            .strip()
        )


        file_data = base64.b64decode(
            encoded_content
        )


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


        # Make sure every required column exists
        for column in DEFAULT_COLUMNS:

            if column not in dataframe.columns:

                dataframe[column] = ""


        dataframe = dataframe[
            DEFAULT_COLUMNS
        ]


        # Rebuild S.No
        dataframe["S.No"] = range(
            1,
            len(dataframe) + 1
        )


        # Make STOCK numeric
        dataframe["STOCK"] = pd.to_numeric(
            dataframe["STOCK"],
            errors="coerce"
        ).fillna(0).astype(int)


        return dataframe


    except Exception as e:

        st.error(
            f"❌ Error loading inventory:\n\n{e}"
        )

        return pd.DataFrame(
            columns=DEFAULT_COLUMNS
        )


# ==============================================================================
# SAVE INVENTORY TO GITHUB
# ==============================================================================

def save_to_github(dataframe):

    if not GITHUB_TOKEN:

        st.error(
            "❌ GITHUB_TOKEN is missing from Streamlit Secrets."
        )

        return False


    try:

        # ----------------------------------------------------------------------
        # Create Excel file in memory
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
        # Convert to Base64
        # ----------------------------------------------------------------------

        encoded_content = base64.b64encode(
            buffer.getvalue()
        ).decode("utf-8")


        # ----------------------------------------------------------------------
        # Get current SHA
        # ----------------------------------------------------------------------

        response = requests.get(
            API_URL,
            headers=get_github_headers(),
            timeout=30
        )


        if response.status_code != 200:

            st.error(
                f"❌ Could not access Inventory.xlsx.\n\n"
                f"HTTP Status: {response.status_code}\n\n"
                f"{response.text[:500]}"
            )

            return False


        github_file = response.json()

        sha = github_file.get("sha")


        if not sha:

            st.error(
                "❌ GitHub did not return the file SHA."
            )

            return False


        # ----------------------------------------------------------------------
        # Upload
        # ----------------------------------------------------------------------

        payload = {

            "message":
                "Update inventory via Streamlit",

            "content":
                encoded_content,

            "sha":
                sha
        }


        upload_response = requests.put(
            API_URL,
            headers=get_github_headers(),
            json=payload,
            timeout=30
        )


        if upload_response.status_code in [200, 201]:

            return True


        st.error(
            f"❌ GitHub save failed.\n\n"
            f"HTTP Status: {upload_response.status_code}\n\n"
            f"{upload_response.text[:500]}"
        )

        return False


    except Exception as e:

        st.error(
            f"❌ GitHub synchronization error:\n\n{e}"
        )

        return False


# ==============================================================================
# LOAD DATA
# ==============================================================================

if (
    st.session_state.inventory_df is None
):

    with st.spinner(
        "Loading inventory from GitHub..."
    ):

        st.session_state.inventory_df = (
            load_from_github()
        )


df = st.session_state.inventory_df


# ==============================================================================
# CLEAN DATA
# ==============================================================================

for column in DEFAULT_COLUMNS:

    if column not in df.columns:

        df[column] = ""


df = df[
    DEFAULT_COLUMNS
]


df["S.No"] = range(
    1,
    len(df) + 1
)


df["STOCK"] = pd.to_numeric(
    df["STOCK"],
    errors="coerce"
).fillna(0).astype(int)


st.session_state.inventory_df = df




# ==============================================================================
# SEARCH
# ==============================================================================

st.markdown("---")

st.subheader("🔎 Search Inventory")


search_col1, search_col2 = st.columns(
    [4, 1]
)


with search_col1:

    search_query = st.text_input(
        "Search",
        placeholder="Search equipment...",
        label_visibility="collapsed"
    )


with search_col2:

    reset = st.button(
        "🔄 Reset",
        use_container_width=True
    )


if reset:

    st.rerun()


# ==============================================================================
# FILTER
# ==============================================================================

if search_query:

    display_df = df[
        df["EQUIPMENT"]
        .astype(str)
        .str.contains(
            search_query,
            case=False,
            na=False
        )
    ]

else:

    display_df = df


# ==============================================================================
# INVENTORY TABLE
# ==============================================================================

# ==============================================================================
# INVENTORY TABLE - EDITABLE
# ==============================================================================
# ==============================================================================
# INVENTORY TABLE - EDITABLE
# ==============================================================================

st.markdown("---")

st.subheader("📋 Current Stock Inventory")

# Create editable copy
editable_df = display_df.copy()

# Editable inventory table
edited_df = st.data_editor(
    editable_df,
    use_container_width=True,
    hide_index=True,
    num_rows="fixed",

    column_config={

        "S.No": st.column_config.NumberColumn(
            "S.No",
            disabled=True,
            width="small"
        ),

        "EQUIPMENT": st.column_config.TextColumn(
            "Equipment",
            required=True
        ),

        "LASERAX PROJECT No. - Part NO": st.column_config.TextColumn(
            "Project / Part No."
        ),

        "STOCK": st.column_config.NumberColumn(
            "Stock",
            min_value=0,
            step=1
        ),

        "LOCATION": st.column_config.TextColumn(
            "Location"
        ),

        "REMARKS": st.column_config.TextColumn(
            "Remarks"
        ),

        "PROCUREMENT LINK": st.column_config.LinkColumn(
            "Procurement Link"
        )
    },

    disabled=["S.No"],

    key=f"inventory_editor_{search_query}"
)


# ==============================================================================
# SAVE TABLE CHANGES
# ==============================================================================

if st.button(
    "💾 Save Table Changes",
    type="primary",
    use_container_width=True
):

    updated_df = st.session_state.inventory_df.copy()

    # --------------------------------------------------------------------------
    # Update rows from the edited table
    # --------------------------------------------------------------------------

    for _, edited_row in edited_df.iterrows():

        # Find original row using S.No
        matching_indices = updated_df.index[
            updated_df["S.No"] == edited_row["S.No"]
        ]

        if len(matching_indices) == 0:
            st.error(
                f"❌ Could not find inventory row with "
                f"S.No {edited_row['S.No']}."
            )
            st.stop()

        original_index = matching_indices[0]

        # ----------------------------------------------------------------------
        # Equipment
        # ----------------------------------------------------------------------

        if pd.isna(edited_row["EQUIPMENT"]):

            st.error(
                f"❌ Equipment name cannot be empty for row "
                f"{edited_row['S.No']}."
            )

            st.stop()

        updated_df.at[
            original_index,
            "EQUIPMENT"
        ] = str(
            edited_row["EQUIPMENT"]
        ).strip()

        # ----------------------------------------------------------------------
        # Project / Part No.
        # ----------------------------------------------------------------------

        updated_df.at[
            original_index,
            "LASERAX PROJECT No. - Part NO"
        ] = (
            ""
            if pd.isna(
                edited_row["LASERAX PROJECT No. - Part NO"]
            )
            else str(
                edited_row["LASERAX PROJECT No. - Part NO"]
            ).strip()
        )

        # ----------------------------------------------------------------------
        # Stock
        # ----------------------------------------------------------------------

        stock_value = pd.to_numeric(
            edited_row["STOCK"],
            errors="coerce"
        )

        if pd.isna(stock_value):
            stock_value = 0

        updated_df.at[
            original_index,
            "STOCK"
        ] = max(
            0,
            int(stock_value)
        )

        # ----------------------------------------------------------------------
        # Location
        # ----------------------------------------------------------------------

        updated_df.at[
            original_index,
            "LOCATION"
        ] = (
            ""
            if pd.isna(edited_row["LOCATION"])
            else str(
                edited_row["LOCATION"]
            ).strip()
        )

        # ----------------------------------------------------------------------
        # Remarks
        # ----------------------------------------------------------------------

        updated_df.at[
            original_index,
            "REMARKS"
        ] = (
            ""
            if pd.isna(edited_row["REMARKS"])
            else str(
                edited_row["REMARKS"]
            ).strip()
        )

        # ----------------------------------------------------------------------
        # Procurement Link
        # ----------------------------------------------------------------------

        updated_df.at[
            original_index,
            "PROCUREMENT LINK"
        ] = (
            ""
            if pd.isna(edited_row["PROCUREMENT LINK"])
            else str(
                edited_row["PROCUREMENT LINK"]
            ).strip()
        )

    # --------------------------------------------------------------------------
    # Rebuild serial numbers
    # --------------------------------------------------------------------------

    updated_df["S.No"] = range(
        1,
        len(updated_df) + 1
    )

    # --------------------------------------------------------------------------
    # Save to GitHub
    # --------------------------------------------------------------------------

    with st.spinner(
        "Saving changes to GitHub..."
    ):

        if save_to_github(updated_df):

            st.session_state.inventory_df = updated_df

            st.success(
                "✅ Inventory table changes saved successfully!"
            )

            st.rerun()

# ==============================================================================
# ADD INVENTORY ITEM
# ==============================================================================

st.markdown("---")

st.subheader("➕ Add Inventory Item")

st.info(
    "Add a new inventory record below. "
    "The new item will be added to the Excel inventory and synchronized with GitHub."
)


add_col1, add_col2, add_col3 = st.columns(3)


with add_col1:

    add_eq = st.text_input(
        "Equipment Name",
        key="add_equipment"
    )


with add_col2:

    add_proj = st.text_input(
        "Project / Part No.",
        key="add_project"
    )


with add_col3:

    add_stock = st.number_input(
        "Stock",
        min_value=0,
        value=0,
        step=1,
        key="add_stock"
    )


add_col4, add_col5, add_col6 = st.columns(3)


with add_col4:

    add_loc = st.text_input(
        "Location",
        key="add_location"
    )


with add_col5:

    add_rem = st.text_input(
        "Remarks",
        key="add_remarks"
    )


with add_col6:

    add_link = st.text_input(
        "Procurement Link",
        key="add_link"
    )


if st.button(
    "💾 Add Inventory Item",
    type="primary",
    use_container_width=True
):

    if not add_eq.strip():

        st.error(
            "❌ Equipment Name is required."
        )

    else:

        # Create new inventory row
        new_row = pd.DataFrame([{

            "S.No":
                len(df) + 1,

            "EQUIPMENT":
                add_eq.strip(),

            "LASERAX PROJECT No. - Part NO":
                add_proj.strip(),

            "STOCK":
                int(add_stock),

            "LOCATION":
                add_loc.strip(),

            "REMARKS":
                add_rem.strip(),

            "PROCUREMENT LINK":
                add_link.strip()

        }])


        # Add new row to existing inventory
        updated_df = pd.concat(
            [
                st.session_state.inventory_df,
                new_row
            ],
            ignore_index=True
        )


        # Rebuild S.No
        updated_df["S.No"] = range(
            1,
            len(updated_df) + 1
        )


        # Make stock numeric
        updated_df["STOCK"] = pd.to_numeric(
            updated_df["STOCK"],
            errors="coerce"
        ).fillna(0).astype(int)


        # Save to GitHub
        with st.spinner(
            "Adding inventory item..."
        ):

            if save_to_github(
                updated_df
            ):

                st.session_state.inventory_df = (
                    updated_df
                )

                st.success(
                    "✅ New inventory item added successfully!"
                )

                st.rerun()


# ==============================================================================
# REMOVE INVENTORY ITEM
# ==============================================================================

st.markdown("---")

st.subheader("➖ Remove Inventory Item")

st.warning(
    "⚠️ Removing an item will permanently delete it from the inventory."
)


if len(df) == 0:

    st.info(
        "No inventory records available."
    )

else:

    # --------------------------------------------------------------------------
    # Select inventory item
    # --------------------------------------------------------------------------

    remove_options = [

        f"Row {row['S.No']}: {row['EQUIPMENT']}"

        for _, row in df.iterrows()

    ]


    remove_selected = st.selectbox(
        "Select inventory item to remove",
        remove_options,
        key="remove_inventory_select"
    )


    # Get selected S.No
    remove_sno = int(
        remove_selected
        .split(": ")[0]
        .replace("Row ", "")
        .strip()
    )


    # Find selected row
    remove_matching = df[
        df["S.No"] == remove_sno
    ]


    if len(remove_matching) > 0:

        remove_row_idx = remove_matching.index[0]

        remove_current = df.loc[
            remove_row_idx
        ]


        # Show selected item information

        remove_col1, remove_col2, remove_col3 = st.columns(3)


        with remove_col1:

            st.metric(
                "Equipment",
                str(
                    remove_current["EQUIPMENT"]
                )
            )


        with remove_col2:

            st.metric(
                "Stock",
                int(
                    remove_current["STOCK"]
                )
            )


        with remove_col3:

            st.metric(
                "Location",
                str(
                    remove_current["LOCATION"]
                )
            )


        # ----------------------------------------------------------------------
        # Remove button
        # ----------------------------------------------------------------------

        confirm_key = (
            f"confirm_remove_{remove_sno}"
        )


        if st.button(
            "🗑️ Remove Selected Inventory",
            type="primary",
            use_container_width=True
        ):

            st.session_state[
                confirm_key
            ] = True


        # ----------------------------------------------------------------------
        # Confirmation
        # ----------------------------------------------------------------------

        if st.session_state.get(
            confirm_key,
            False
        ):

            st.error(
                f"⚠️ You are about to remove "
                f"**{remove_current['EQUIPMENT']}**."
            )


            confirm_col1, confirm_col2 = st.columns(2)


            with confirm_col1:

                if st.button(
                    "✅ Yes, Remove Item",
                    key=f"confirm_yes_{remove_sno}",
                    type="primary",
                    use_container_width=True
                ):

                    updated_df = (
                        st.session_state.inventory_df
                        .drop(index=remove_row_idx)
                        .reset_index(drop=True)
                    )


                    # Rebuild S.No
                    updated_df["S.No"] = range(
                        1,
                        len(updated_df) + 1
                    )


                    # Make stock numeric
                    updated_df["STOCK"] = pd.to_numeric(
                        updated_df["STOCK"],
                        errors="coerce"
                    ).fillna(0).astype(int)


                    # Save to GitHub
                    with st.spinner(
                        "Removing inventory item..."
                    ):

                        if save_to_github(
                            updated_df
                        ):

                            st.session_state.inventory_df = (
                                updated_df
                            )

                            st.session_state[
                                confirm_key
                            ] = False

                            st.success(
                                "🗑️ Inventory item removed successfully."
                            )

                            st.rerun()


            with confirm_col2:

                if st.button(
                    "❌ Cancel",
                    key=f"confirm_cancel_{remove_sno}",
                    use_container_width=True
                ):

                    st.session_state[
                        confirm_key
                    ] = False

                    st.rerun()


# ==============================================================================
# FOOTER
# ==============================================================================


st.caption(
    "☁️ Inventory data is synchronized with GitHub."
)

st.caption(
    f"Total inventory records: {len(df)}"
)

