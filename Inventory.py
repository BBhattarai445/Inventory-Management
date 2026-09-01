
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

GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")


# ==============================================================================
# LOGIN CONFIGURATION
#
# Put these values in Streamlit Secrets.
#
# ADMIN_USERNAME = "admin"
# ADMIN_PASSWORD = "your-admin-password"
# USER_USERNAME = "user"
# USER_PASSWORD = "your-user-password"
#
# ==============================================================================

ADMIN_USERNAME = st.secrets.get("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "")

USER_USERNAME = st.secrets.get("USER_USERNAME", "user")
USER_PASSWORD = st.secrets.get("USER_PASSWORD", "")


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


# ==============================================================================
# LOGIN FUNCTION
# ==============================================================================

def login_page():

    st.title("🏭 Laserax Inventory Manager")

    st.markdown("---")

    st.subheader("🔐 Login")

    login_col1, login_col2, login_col3 = st.columns(
        [1, 2, 1]
    )

    with login_col2:

        username = st.text_input(
            "Username",
            placeholder="Enter username"
        )

        password = st.text_input(
            "Password",
            type="password",
            placeholder="Enter password"
        )

        login_clicked = st.button(
            "🔐 Login",
            type="primary",
            use_container_width=True
        )

        if login_clicked:

            # --------------------------------------------------------------
            # ADMIN LOGIN
            # --------------------------------------------------------------

            if (
                username == ADMIN_USERNAME
                and password == ADMIN_PASSWORD
                and ADMIN_PASSWORD
            ):

                st.session_state.authenticated = True
                st.session_state.role = "admin"
                st.session_state.username = username

                st.rerun()


            # --------------------------------------------------------------
            # USER LOGIN
            # --------------------------------------------------------------

            elif (
                username == USER_USERNAME
                and password == USER_PASSWORD
                and USER_PASSWORD
            ):

                st.session_state.authenticated = True
                st.session_state.role = "user"
                st.session_state.username = username

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
# SHOW LOGIN PAGE IF NOT AUTHENTICATED
# ==============================================================================

if not st.session_state.authenticated:

    login_page()
    st.stop()


# ==============================================================================
# CURRENT USER INFORMATION
# ==============================================================================

is_admin = st.session_state.role == "admin"
is_user = st.session_state.role == "user"


# ==============================================================================
# SIDEBAR
# ==============================================================================

with st.sidebar:

    st.markdown("## 👤 Account")

    if is_admin:

        st.success(
            f"👑 Admin\n\n"
            f"Logged in as: **{st.session_state.username}**"
        )

    else:

        st.info(
            f"👤 Inventory User\n\n"
            f"Logged in as: **{st.session_state.username}**"
        )


    st.markdown("---")


    # --------------------------------------------------------------------------
    # LOGOUT
    # --------------------------------------------------------------------------

    if st.button(
        "🚪 Logout",
        use_container_width=True
    ):

        st.session_state.authenticated = False
        st.session_state.role = None
        st.session_state.username = None

        if "inventory_df" in st.session_state:
            del st.session_state.inventory_df

        st.rerun()


    # --------------------------------------------------------------------------
    # ADMIN-ONLY GITHUB SYNC
    # --------------------------------------------------------------------------

    if is_admin:

        st.markdown("---")
        st.markdown("### 👑 Admin Controls")

        sync_clicked = st.button(
            "🔄 Sync Live GitHub Data",
            use_container_width=True
        )

    else:

        sync_clicked = False


# ==============================================================================
# PAGE TITLE
# ==============================================================================

st.title("🏭 Laserax Inventory Manager")

if is_admin:

    st.caption(
        "👑 Administrator Mode — Full Inventory Management"
    )

else:

    st.caption(
        "👤 Inventory User Mode — Existing Inventory Editing Only"
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
# LOAD FROM GITHUB
# ==============================================================================

def load_from_github():

    if not GITHUB_TOKEN:

        st.error(
            "❌ Missing GITHUB_TOKEN in Streamlit Secrets."
        )

        return pd.DataFrame(
            columns=DEFAULT_COLUMNS
        )


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

            return pd.DataFrame(
                columns=DEFAULT_COLUMNS
            )


        github_data = response.json()


        if "content" not in github_data:

            st.error(
                "❌ GitHub API response did not contain file content."
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


        dataframe.columns = (
            dataframe.columns
            .astype(str)
            .str.strip()
        )


        # Make sure all columns exist

        for column in DEFAULT_COLUMNS:

            if column not in dataframe.columns:

                dataframe[column] = ""


        dataframe = dataframe[
            DEFAULT_COLUMNS
        ]


        dataframe["S.No"] = range(
            1,
            len(dataframe) + 1
        )


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
# SAVE TO GITHUB
# ==============================================================================

def save_to_github(dataframe):

    if not GITHUB_TOKEN:

        st.error(
            "❌ Missing GITHUB_TOKEN in Streamlit Secrets."
        )

        return False


    try:

        # ----------------------------------------------------------------------
        # Create Excel in memory
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
        # Encode Excel as Base64
        # ----------------------------------------------------------------------

        content_encoded = base64.b64encode(
            buffer.getvalue()
        ).decode("utf-8")


        # ----------------------------------------------------------------------
        # Get current GitHub SHA
        # ----------------------------------------------------------------------

        headers = get_github_headers()


        get_response = requests.get(
            API_URL,
            headers=headers,
            timeout=30
        )


        if get_response.status_code != 200:

            st.error(
                f"❌ Could not access Inventory.xlsx on GitHub.\n\n"
                f"HTTP Status: {get_response.status_code}\n"
                f"Response: {get_response.text[:500]}"
            )

            return False


        github_file = get_response.json()

        sha = github_file.get("sha")


        if not sha:

            st.error(
                "❌ GitHub did not return the file SHA."
            )

            return False


        # ----------------------------------------------------------------------
        # Upload updated Excel
        # ----------------------------------------------------------------------

        payload = {

            "message":
                "Update inventory via Laserax Inventory Manager",

            "content":
                content_encoded,

            "sha":
                sha
        }


        put_response = requests.put(
            API_URL,
            headers=headers,
            json=payload,
            timeout=30
        )


        if put_response.status_code in [200, 201]:

            st.toast(
                "☁️ Inventory saved to GitHub!",
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
# INITIAL LOAD
# ==============================================================================

if (
    "inventory_df" not in st.session_state
    or sync_clicked
):

    with st.spinner(
        "Loading inventory from GitHub..."
    ):

        st.session_state.inventory_df = (
            load_from_github()
        )


# ==============================================================================
# MAKE SURE DATAFRAME EXISTS
# ==============================================================================

if "inventory_df" not in st.session_state:

    st.session_state.inventory_df = (
        pd.DataFrame(
            columns=DEFAULT_COLUMNS
        )
    )


df = st.session_state.inventory_df


# ==============================================================================
# CLEAN DATAFRAME
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
    [3, 1]
)


with search_col1:

    search_query = st.text_input(
        "Search Equipment",
        placeholder="Type equipment name...",
        label_visibility="collapsed"
    )


with search_col2:

    reset_clicked = st.button(
        "🔄 Reset View",
        use_container_width=True
    )


if reset_clicked:

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

st.markdown("---")

st.subheader("📋 Current Stock Inventory")


st.dataframe(
    display_df,
    use_container_width=True,
    hide_index=True
)


st.caption(
    f"Showing {len(display_df)} of {len(df)} records."
)


# ==============================================================================
# ADMIN ONLY — ADD NEW ITEM
# ==============================================================================

if is_admin:

    st.markdown("---")

    st.subheader(
        "👑 ➕ Add New Inventory Line Item"
    )


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
        "💾 Add New Inventory Item",
        type="primary",
        use_container_width=True
    ):

        if not add_eq.strip():

            st.error(
                "❌ Equipment Name cannot be empty."
            )

        else:

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


            updated_df = pd.concat(
                [df, new_row],
                ignore_index=True
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

                st.success(
                    "✅ New inventory item added."
                )

                st.rerun()


# ==============================================================================
# EDIT EXISTING INVENTORY
# ==============================================================================

st.markdown("---")

if is_admin:

    st.subheader(
        "✏️ 👑 Manage Existing Inventory"
    )

else:

    st.subheader(
        "✏️ Edit Existing Inventory"
    )


if len(df) == 0:

    st.info(
        "There are currently no inventory records."
    )

else:

    select_options = [

        f"Row {row['S.No']}: {row['EQUIPMENT']}"

        for _, row in df.iterrows()
    ]


    selected_option = st.selectbox(
        "Choose inventory record",
        select_options
    )


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
                "❌ Selected record could not be found."
            )

            row_idx = None

        else:

            row_idx = matching_rows.index[0]

            current_row = df.loc[row_idx]


    except Exception as e:

        st.error(
            f"❌ Selection error: {e}"
        )

        row_idx = None


    # ==========================================================================
    # EDIT FIELDS
    # ==========================================================================

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


        # ======================================================================
        # ACTION BUTTONS
        # ======================================================================

        if is_admin:

            action_col1, action_col2, action_col3, action_col4 = st.columns(4)

        else:

            action_col1, action_col2, action_col3 = st.columns(3)


        # ======================================================================
        # INCREASE STOCK — ADMIN AND USER
        # ======================================================================

        with action_col1:

            if st.button(
                "➕ Increase Stock (+1)",
                use_container_width=True
            ):

                updated_df = (
                    st.session_state.inventory_df.copy()
                )


                current_qty = int(
                    updated_df.at[
                        row_idx,
                        "STOCK"
                    ]
                )


                updated_df.at[
                    row_idx,
                    "STOCK"
                ] = current_qty + 1


                if save_to_github(
                    updated_df
                ):

                    st.session_state.inventory_df = (
                        updated_df
                    )

                    st.rerun()


        # ======================================================================
        # DECREASE STOCK — ADMIN AND USER
        # ======================================================================

        with action_col2:

            if st.button(
                "➖ Decrease Stock (-1)",
                use_container_width=True
            ):

                updated_df = (
                    st.session_state.inventory_df.copy()
                )


                current_qty = int(
                    updated_df.at[
                        row_idx,
                        "STOCK"
                    ]
                )


                if current_qty > 0:

                    updated_df.at[
                        row_idx,
                        "STOCK"
                    ] = current_qty - 1


                    if save_to_github(
                        updated_df
                    ):

                        st.session_state.inventory_df = (
                            updated_df
                        )

                        st.rerun()

                else:

                    st.warning(
                        "⚠️ Stock is already 0."
                    )


        # ======================================================================
        # SAVE CHANGES — ADMIN AND USER
        # ======================================================================

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


                    if save_to_github(
                        updated_df
                    ):

                        st.session_state.inventory_df = (
                            updated_df
                        )

                        st.success(
                            "✅ Inventory record updated."
                        )

                        st.rerun()


        # ======================================================================
        # DELETE — ADMIN ONLY
        # ======================================================================

        if is_admin:

            with action_col4:

                if st.button(
                    "🗑️ Delete Item",
                    use_container_width=True
                ):

                    # Confirmation checkbox
                    confirm_delete = st.checkbox(
                        "I confirm I want to delete this item.",
                        key=f"confirm_delete_{selected_sno}"
                    )


                    if confirm_delete:

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

                            st.success(
                                "🗑️ Inventory item deleted."
                            )

                            st.rerun()

                    else:

                        st.warning(
                            "⚠️ Please confirm deletion first."
                        )


# ==============================================================================
# USER PERMISSION NOTICE
# ==============================================================================

if not is_admin:

    st.markdown("---")

    st.info(
        "👤 **User access:** You can edit existing inventory records "
        "and adjust stock quantities. Adding, deleting, and administrative "
        "controls are restricted to the administrator."
    )


# ==============================================================================
# ADMIN STATUS
# ==============================================================================

if is_admin:

    st.markdown("---")

    st.success(
        "👑 Administrator access enabled — full inventory management available."
    )


# ==============================================================================
# FOOTER
# ==============================================================================

st.markdown("---")

st.caption(
    "☁️ Inventory data is synchronized with the GitHub repository."
)

st.caption(
    f"Total inventory records: "
    f"{len(st.session_state.inventory_df)}"
)
```

