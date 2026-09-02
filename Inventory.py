
import streamlit as st
import pandas as pd
import io
import base64
import requests


# ==============================================================================
# PAGE CONFIG
# ==============================================================================

st.set_page_config(
    page_title="Laserax Inventory Manager",
    page_icon= "🏭",
    layout="wide"
)


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

ADMIN_USERNAME = st.secrets.get(
    "ADMIN_USERNAME",
    "admin"
)

ADMIN_PASSWORD = st.secrets.get(
    "ADMIN_PASSWORD",
    ""
)

USER_USERNAME = st.secrets.get(
    "USER_USERNAME",
    "user"
)

USER_PASSWORD = st.secrets.get(
    "USER_PASSWORD",
    ""
)


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


def login_page(): 
    st.title("🏭 Laserax Inventory Manager")

    st.markdown("---")

    st.subheader("🔐 Login")

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:

        username = st.text_input(
            "Username",
            placeholder="Enter username"
        )

        password = st.text_input(
            "Password",
            type="password",
            placeholder="Enter password"
        )

        login = st.button(
            "🔐 Login",
            type="primary",
            use_container_width=True
        )

        if login:

            # ADMIN LOGIN
            if (
                username == ADMIN_USERNAME
                and password == ADMIN_PASSWORD
                and ADMIN_PASSWORD
            ):

                st.session_state.authenticated = True
                st.session_state.role = "admin"
                st.session_state.username = username
                st.session_state.inventory_df = None

                st.rerun()


            # USER LOGIN
            elif (
                username == USER_USERNAME
                and password == USER_PASSWORD
                and USER_PASSWORD
            ):

                st.session_state.authenticated = True
                st.session_state.role = "user"
                st.session_state.username = username
                st.session_state.inventory_df = None

                st.rerun()


            else:

                st.error(
                    "❌ Invalid username or password."
                )

    st.markdown("---")

    st.caption(
        "Authorized users only."
    )


# ==============================================================================
# STOP HERE IF NOT LOGGED IN
# ==============================================================================

if not st.session_state.authenticated:

    login_page()
    st.stop()


# ==============================================================================
# ROLE
# ==============================================================================

is_admin = (
    st.session_state.role == "admin"
)

is_user = (
    st.session_state.role == "user"
)

if not is_admin:
    st.markdown(
        """
        <style>
        [data-testid="stAppDeployButton"] {
            display: none !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )


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
# SIDEBAR
# ==============================================================================

with st.sidebar:

    st.markdown("## 👤 Account")

    if is_admin:
        st.success(
            f"👑 **Administrator**\n\n"
            f"{st.session_state.username}"
        )
    else:
        st.info(
            f"👤 **Inventory User**\n\n"
            f"{st.session_state.username}"
        )

    st.markdown("---")

    # ======================================================================
    # ADMIN ONLY — MANAGE APP
    # ======================================================================

    if is_admin:

        st.markdown("## ⚙️ Manage App")

        st.caption(
            "Administrator-only application controls."
        )


        st.markdown(
            """
            **Admin permissions**
            
            • Application management  
            • Inventory management  
            • GitHub synchronization  
            • System controls
            """
        )


        st.markdown("---")


        if st.button(
            "🔄 Sync GitHub",
            use_container_width=True
        ):

            with st.spinner(
                "Syncing with GitHub..."
            ):

                st.session_state.inventory_df = (
                    load_from_github()
                )

            st.success(
                "✅ Inventory synchronized."
            )

            st.rerun()


    # ======================================================================
    # USER — NO MANAGE APP SECTION
    # ======================================================================

    else:

        st.markdown("## 📦 Inventory")

        st.caption(
            "Inventory management access."
        )


        if st.button(
            "🔄 Sync GitHub",
            use_container_width=True
        ):

            with st.spinner(
                "Syncing with GitHub..."
            ):

                st.session_state.inventory_df = (
                    load_from_github()
                )

            st.success(
                "✅ Inventory synchronized."
            )

            st.rerun()


    st.markdown("---")


    # ======================================================================
    # LOGOUT
    # ======================================================================

    if st.button(
        "🚪 Logout",
        use_container_width=True
    ):

        st.session_state.authenticated = False
        st.session_state.role = None
        st.session_state.username = None
        st.session_state.inventory_df = None

        st.rerun()


# ==============================================================================
# MAIN TITLE
# ==============================================================================

st.title("🏭 Laserax Inventory Manager")


if is_admin:

    st.caption(
        "👑 Administrator Portal"
    )

else:

    st.caption(
        "👤 Inventory User Portal"
    )


# ==============================================================================
# ADMIN-ONLY MANAGE APP PANEL
# ==============================================================================

if is_admin:

    with st.expander(
        "⚙️ Manage App",
        expanded=False
    ):

        st.subheader(
            "Application Management"
        )

        st.info(
            "This section is available only to administrators."
        )


        manage_col1, manage_col2, manage_col3 = st.columns(3)


        with manage_col1:

            st.metric(
                "Inventory Records",
                len(df)
            )


        with manage_col2:

            total_stock = int(
                df["STOCK"].sum()
            )

            st.metric(
                "Total Stock",
                total_stock
            )


        with manage_col3:

            unique_equipment = (
                df["EQUIPMENT"]
                .astype(str)
                .nunique()
            )

            st.metric(
                "Equipment Types",
                unique_equipment
            )


        st.markdown("---")

        st.write(
            "👑 Administrator access includes "
            "application-level controls."
        )


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

st.markdown("---")

st.subheader("📋 Current Stock Inventory")

st.info(
    "✏️ You can edit Stock, Location, Remarks, Project/Part No., "
    "and other fields directly in the table. Click Save Changes when finished."
)

# Create editable copy
editable_df = display_df.copy()

# Keep original index so we know which rows were changed
editable_df["_original_index"] = editable_df.index

# Editable inventory table
edited_df = st.data_editor(
    editable_df.drop(columns=["_original_index"]),
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

    key="inventory_editor"
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

    for position, edited_row in edited_df.iterrows():

        # Get corresponding original row
        original_index = display_df.index[position]

        # Equipment
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


        # Project / Part No.
        updated_df.at[
            original_index,
            "LASERAX PROJECT No. - Part NO"
        ] = (
            ""
            if pd.isna(
                edited_row[
                    "LASERAX PROJECT No. - Part NO"
                ]
            )
            else str(
                edited_row[
                    "LASERAX PROJECT No. - Part NO"
                ]
            ).strip()
        )


        # Stock
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


        # Location
        updated_df.at[
            original_index,
            "LOCATION"
        ] = (
            ""
            if pd.isna(
                edited_row["LOCATION"]
            )
            else str(
                edited_row["LOCATION"]
            ).strip()
        )


        # Remarks
        updated_df.at[
            original_index,
            "REMARKS"
        ] = (
            ""
            if pd.isna(
                edited_row["REMARKS"]
            )
            else str(
                edited_row["REMARKS"]
            ).strip()
        )


        # Procurement Link
        updated_df.at[
            original_index,
            "PROCUREMENT LINK"
        ] = (
            ""
            if pd.isna(
                edited_row["PROCUREMENT LINK"]
            )
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

        if save_to_github(
            updated_df
        ):

            st.session_state.inventory_df = (
                updated_df
            )

            st.success(
                "✅ Inventory table changes saved successfully!"
            )

            st.rerun()


st.caption(
    f"Showing {len(display_df)} of {len(df)} records."
)


# ==============================================================================
# EDIT / DELETE INVENTORY
# ADMIN + USER
# ==============================================================================

st.markdown("---")

st.subheader("✏️ Manage Inventory")


if len(df) == 0:

    st.info(
        "No inventory records available."
    )

else:

    options = [

        f"Row {row['S.No']}: {row['EQUIPMENT']}"

        for _, row in df.iterrows()
    ]


    selected = st.selectbox(
        "Select inventory record",
        options
    )


    selected_sno = int(
        selected
        .split(": ")[0]
        .replace("Row ", "")
        .strip()
    )


    matching = df[
        df["S.No"] == selected_sno
    ]


    if len(matching) > 0:

        row_idx = matching.index[0]

        current = df.loc[row_idx]


        # ==================================================================
        # EDIT FIELDS
        # ==================================================================

        edit_col1, edit_col2, edit_col3 = st.columns(3)


        with edit_col1:

            edit_eq = st.text_input(
                "Equipment Name",
                value=str(
                    current["EQUIPMENT"]
                ),
                key=f"edit_eq_{selected_sno}"
            )


        with edit_col2:

            edit_proj = st.text_input(
                "Project / Part No.",
                value=str(
                    current[
                        "LASERAX PROJECT No. - Part NO"
                    ]
                ),
                key=f"edit_proj_{selected_sno}"
            )


        with edit_col3:

            edit_stock = st.number_input(
                "Stock",
                min_value=0,
                value=int(
                    current["STOCK"]
                ),
                step=1,
                key=f"edit_stock_{selected_sno}"
            )


        edit_col4, edit_col5, edit_col6 = st.columns(3)


        with edit_col4:

            edit_loc = st.text_input(
                "Location",
                value=str(
                    current["LOCATION"]
                ),
                key=f"edit_loc_{selected_sno}"
            )


        with edit_col5:

            edit_rem = st.text_input(
                "Remarks",
                value=str(
                    current["REMARKS"]
                ),
                key=f"edit_rem_{selected_sno}"
            )


        with edit_col6:

            edit_link = st.text_input(
                "Procurement Link",
                value=str(
                    current["PROCUREMENT LINK"]
                ),
                key=f"edit_link_{selected_sno}"
            )


        # ==================================================================
        # ACTION BUTTONS
        # ==================================================================

        action1, action2, action3, action4 = st.columns(4)


        # ==================================================================
        # INCREASE
        # ==================================================================

        with action1:

            if st.button(
                "➕ Increase Stock",
                use_container_width=True
            ):

                updated_df = (
                    st.session_state.inventory_df.copy()
                )


                updated_df.at[
                    row_idx,
                    "STOCK"
                ] = (
                    int(
                        updated_df.at[
                            row_idx,
                            "STOCK"
                        ]
                    ) + 1
                )


                if save_to_github(
                    updated_df
                ):

                    st.session_state.inventory_df = (
                        updated_df
                    )

                    st.success(
                        "✅ Stock increased."
                    )

                    st.rerun()


        # ==================================================================
        # DECREASE
        # ==================================================================

        with action2:

            if st.button(
                "➖ Decrease Stock",
                use_container_width=True
            ):

                updated_df = (
                    st.session_state.inventory_df.copy()
                )


                current_stock = int(
                    updated_df.at[
                        row_idx,
                        "STOCK"
                    ]
                )


                if current_stock > 0:

                    updated_df.at[
                        row_idx,
                        "STOCK"
                    ] = current_stock - 1


                    if save_to_github(
                        updated_df
                    ):

                        st.session_state.inventory_df = (
                            updated_df
                        )

                        st.success(
                            "✅ Stock decreased."
                        )

                        st.rerun()

                else:

                    st.warning(
                        "⚠️ Stock is already 0."
                    )


        # ==================================================================
        # SAVE
        # ==================================================================

        with action3:

            if st.button(
                "💾 Save Changes",
                type="primary",
                use_container_width=True
            ):

                if not edit_eq.strip():

                    st.error(
                        "❌ Equipment Name is required."
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


                    if save_to_github(
                        updated_df
                    ):

                        st.session_state.inventory_df = (
                            updated_df
                        )

                        st.success(
                            "✅ Changes saved."
                        )

                        st.rerun()


        # ==================================================================
        # DELETE
        # ==================================================================

        with action4:

            if st.button(
                "🗑️ Delete Item",
                use_container_width=True
            ):

                st.session_state[
                    f"confirm_delete_{selected_sno}"
                ] = True


            if st.session_state.get(
                f"confirm_delete_{selected_sno}",
                False
            ):

                st.warning(
                    "⚠️ Are you sure you want to delete this item?"
                )


                confirm_col1, confirm_col2 = st.columns(2)


                with confirm_col1:

                    if st.button(
                        "✅ Yes, Delete",
                        key=f"yes_delete_{selected_sno}",
                        use_container_width=True
                    ):

                        updated_df = (
                            st.session_state.inventory_df
                            .drop(index=row_idx)
                            .reset_index(drop=True)
                        )


                        updated_df["S.No"] = range(
                            1,
                            len(updated_df) + 1
                        )


                        if save_to_github(
                            updated_df
                        ):

                            st.session_state.inventory_df = (
                                updated_df
                            )

                            st.session_state[
                                f"confirm_delete_{selected_sno}"
                            ] = False

                            st.success(
                                "🗑️ Item deleted successfully."
                            )

                            st.rerun()


                with confirm_col2:

                    if st.button(
                        "❌ Cancel",
                        key=f"cancel_delete_{selected_sno}",
                        use_container_width=True
                    ):

                        st.session_state[
                            f"confirm_delete_{selected_sno}"
                        ] = False

                        st.rerun()


# ==============================================================================
# FOOTER
# ==============================================================================

st.markdown("---")

if is_admin:

    st.success(
        "👑 Admin Portal — Inventory + Application Management"
    )

else:

    st.info(
        "👤 User Portal — Inventory Management"
    )


st.caption(
    "☁️ Inventory data is synchronized with GitHub."
)

st.caption(
    f"Total inventory records: {len(df)}"
)

