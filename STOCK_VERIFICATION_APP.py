import streamlit as st
import pandas as pd
import easyocr
from PIL import Image
import io
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import format_cell_range, CellFormat, Color  # 🔥 NEW

# ================== GOOGLE SHEET ==================
def connect_sheet():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]

    import json
    creds_dict = st.secrets["gcp_service_account"]

    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)

    return client.open("DAY_TO_DAY_INVENTORY").sheet1

def load_data():
    sheet = connect_sheet()
    data = sheet.get_all_values()

    if not data:
        return pd.DataFrame()

    headers = data[0]
    rows = data[1:] if len(data) > 1 else []

    df = pd.DataFrame(rows, columns=headers)
    return df.fillna("")

def save_data(df):
    if df is None or df.empty:
        st.warning("⚠️ No data to save")
        return

    sheet = connect_sheet()
    df = df.fillna("").astype(str)

    # 🔥 SAVE DATA
    sheet.update([df.columns.values.tolist()] + df.values.tolist())

    # 🔥 APPLY COLORS IN SHEET
    try:
        name_col_index = list(df.columns).index(name_col) + 1  # 1-based index

        for i, row in df.iterrows():
            cell = f"{chr(64 + name_col_index)}{i + 2}"  # Excel style

            if "__status__" in df.columns:
                if row["__status__"] == "updated":
                    format_cell_range(
                        sheet,
                        cell,
                        CellFormat(backgroundColor=Color(0.7, 1, 0.7))  # green
                    )

                elif row["__status__"] == "new":
                    format_cell_range(
                        sheet,
                        cell,
                        CellFormat(backgroundColor=Color(0.7, 0.85, 1))  # blue
                    )
    except:
        pass

# ================== OCR ==================
reader = easyocr.Reader(['en'], gpu=False)

st.title("📦 Stock Verification System")

df_main = load_data()

if df_main is not None and not df_main.empty:

    columns = df_main.columns.tolist()

    def find_column(keys):
        for col in columns:
            for k in keys:
                if k.lower() in col.lower():
                    return col
        return None

    name_col = find_column(["name"])
    barcode_col = find_column(["barcode"])
    mrp_col = find_column(["mrp"])
    verify_col = find_column(["verified"])

    if not verify_col:
        verify_col = "STOCK VERIFIED Y/N"
        df_main[verify_col] = ""

    # ================== PART 1 ==================
    st.subheader("🔍 Scan Product")

    barcode_input = st.text_input("Scan / Enter Barcode")

    if barcode_input:
        match = df_main[df_main[barcode_col].astype(str) == barcode_input.strip()]

        if not match.empty:
            row = match.iloc[0]

            st.success("Product Found")
            st.dataframe(match)

            st.session_state["auto_name"] = row[name_col]
            st.session_state["auto_mrp"] = row[mrp_col] if mrp_col else ""

            if st.button("Mark Verified"):
                st.session_state["backup_df"] = df_main.copy()

                df_main.loc[df_main[barcode_col] == barcode_input.strip(), verify_col] = "Y"

                save_data(df_main)
                st.success("Saved")

        else:
            st.error("Not Found")

    # ================== UNDO VERIFY ==================
    if "backup_df" in st.session_state:
        if st.button("↩️ Undo Last Verify"):
            df_main = st.session_state["backup_df"]
            save_data(df_main)
            st.success("Undo Done")

    # ================== PART 2 ==================
    st.subheader("🔍 Fuzzy Search")

    text_input = st.text_input("Product Name", value=st.session_state.get("auto_name", ""))
    
    text_input = text_input.lower().strip()
    
    user_mrp = st.text_input("MRP", value=st.session_state.get("auto_mrp", ""))

    if text_input:

        from rapidfuzz import fuzz

        df = df_main.copy()
        
        df[name_col] = df[name_col].astype(str).str.lower().str.strip()
        
        results = []

        for _, row in df.iterrows():
            name = str(row[name_col]).lower().strip()

            score = (
                fuzz.token_set_ratio(text_input, name) * 0.4 +
                fuzz.partial_ratio(text_input, name) * 0.3 +
                fuzz.token_sort_ratio(text_input, name) * 0.3
            )

            if user_mrp and mrp_col:
                try:
                    if abs(float(row[mrp_col]) - float(user_mrp)) < 20:
                        score += 20
                except:
                    pass

            results.append((row, score))

        results = sorted(results, key=lambda x: x[1], reverse=True)[:20]

        output = []
        for r, s in results:
            output.append({
                "ITEM NAME": r[name_col],
                "BARCODE": r[barcode_col],
                "MRP": r[mrp_col] if mrp_col else "",
                "MATCH %": round(s, 1)
            })

        st.dataframe(pd.DataFrame(output))

    # ================== MARG UPDATE ==================
    st.subheader("🔄 Marg Update")

    marg_file = st.file_uploader("Upload Marg File", type=["xlsx"])

    if marg_file:

        df_new = pd.read_excel(marg_file).fillna("")

        new_name_col = df_new.columns[0]
        new_barcode_col = df_new.columns[1]
        new_mrp_col = df_new.columns[3] if len(df_new.columns) > 3 else None

        st.session_state["backup_df"] = df_main.copy()

        df_main["__status__"] = ""

        for i, row in df_main.iterrows():
            bc = str(row[barcode_col])
            match = df_new[df_new[new_barcode_col].astype(str) == bc]

            if not match.empty:
                new_val = match.iloc[0][new_name_col]

                if str(row[name_col]) != str(new_val):
                    df_main.at[i, name_col] = new_val
                    df_main.at[i, "__status__"] = "updated"

        existing = set(df_main[barcode_col].astype(str))

        for _, row in df_new.iterrows():
            bc = str(row[new_barcode_col])

            if bc not in existing:
                new_row = {col: "" for col in df_main.columns}

                new_row[name_col] = row[new_name_col]
                new_row[barcode_col] = bc

                if mrp_col and new_mrp_col:
                    new_row[mrp_col] = row[new_mrp_col]

                new_row["__status__"] = "new"

                df_main = pd.concat([df_main, pd.DataFrame([new_row])], ignore_index=True)

        save_data(df_main)

        st.success("Marg Update Applied")

    # ================== UNDO MARG ==================
    if "backup_df" in st.session_state:
        if st.button("Undo Marg Update"):
            df_main = st.session_state["backup_df"]
            save_data(df_main)
            st.success("Reverted")
