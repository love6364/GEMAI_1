
# import streamlit as st
# import pandas as pd
# import io

# st.set_page_config(page_title="💎 Diamond Automation", layout="wide")

# st.title("💎 Diamond Cleaner (Step 1 + 2 + 3)")

# uploaded_file = st.file_uploader("Upload Main Date Excel File", type=["xlsx"])

# # ---------------- STEP 1 ----------------
# def remove_unwanted_columns(df):
#     df.columns = df.columns.str.strip()

#     columns_to_remove = [
#         "Polish", "Sym.", "Flu. Int.", "Tab %", "Dep %", "Cut", "Origin",
#         "List Price", "% Off",
#         "Price A%", "Price A", "Price B%", "Price B", "%RP/Cost",
#         "Rect Cost", "Other Cost", "P&L", "P&&L", "S. Qlty.",
#         "General Note", "Private Note",
#         "CN", "SN", "CW", "SW", "Milky", "Im", "Md", "Im Md",
#         "Itemserial", "Sp", "Price", "Cts",
#         "Lab2", "Cert2", "Price/Cts1", "Price/Cts2",
#         "Price B", "6", "Im Md Itemserial Sp Price Cts"
#     ]

#     df = df.drop(columns=[col for col in columns_to_remove if col in df.columns], errors='ignore')

#     return df


# # ---------------- STEP 2 ----------------
# def filter_lab(df):
#     df["Lab"] = df["Lab"].astype(str).str.strip().str.upper()
#     allowed_labs = ["GIA", "IGI", "GCAL"]
#     return df[df["Lab"].isin(allowed_labs)]


# # ---------------- STEP 3 ----------------
# def fill_quality(df):
#     df.columns = df.columns.str.strip()

#     df["Quality"] = df["Quality"].fillna("").astype(str).str.strip()
#     df["Rapnet Note"] = df["Rapnet Note"].fillna("").astype(str).str.upper()

#     # Safety check
#     if "Lot #" not in df.columns:
#         raise ValueError("❌ 'Lot #' column not found")

#     # Mapping Lot # → Rapnet Note
#     rapnet_map = df.set_index("Lot #")["Rapnet Note"].to_dict()

#     def update_quality(row):
#         if row["Quality"] == "":
#             rap_val = rapnet_map.get(row["Lot #"], "")
#             if "CVD" in rap_val:
#                 return "CVD"
#             elif "HPHT" in rap_val:
#                 return "HPHT"
#         return row["Quality"]

#     df["Quality"] = df.apply(update_quality, axis=1)

#     # Remove Rapnet Note after use
#     df = df.drop(columns=["Rapnet Note"], errors="ignore")

#     return df


# # ---------------- MAIN ----------------
# if uploaded_file:
#     df = pd.read_excel(uploaded_file)

#     st.subheader("📊 Original Data")
#     st.dataframe(df.head())

#     # Required columns check
#     required_cols = ["Lot #", "Lab", "Quality", "Rapnet Note"]
#     missing = [col for col in required_cols if col not in df.columns]

#     if missing:
#         st.error(f"❌ Missing columns: {missing}")
#         st.stop()

#     if st.button("🚀 Process File"):
#         try:
#             # Apply all steps
#             df = remove_unwanted_columns(df)
#             df = filter_lab(df)
#             df = fill_quality(df)

#             st.subheader("✅ Final Processed Data")
#             st.dataframe(df.head())

#             # Download file
#             output = io.BytesIO()
#             with pd.ExcelWriter(output, engine='openpyxl') as writer:
#                 df.to_excel(writer, index=False)

#             st.download_button(
#                 label="📥 Download Final File",
#                 data=output.getvalue(),
#                 file_name="final_cleaned_file.xlsx",
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#             )

#         except Exception as e:
#             st.error(f"❌ Error: {e}")


# import streamlit as st
# import pandas as pd
# import io

# st.set_page_config(page_title="💎 Diamond Automation", layout="wide")

# st.title("💎 Diamond Automation (Step 1 → 4)")

# # Upload both files
# main_file = st.file_uploader("Upload Main Date File", type=["xlsx"])
# labgrown_file = st.file_uploader("Upload Lab Grown File", type=["xlsx"])


# # ---------------- STEP 1 ----------------
# def remove_unwanted_columns(df):
#     df.columns = df.columns.str.strip()

#     columns_to_remove = [
#         "Polish", "Sym.", "Flu. Int.", "Tab %", "Dep %", "Cut", "Origin",
#         "List Price", "% Off",
#         "Price A%", "Price A", "Price B%", "Price B", "%RP/Cost",
#         "Rect Cost", "Other Cost", "P&L", "P&&L", "S. Qlty.",
#         "General Note", "Private Note",
#         "CN", "SN", "CW", "SW", "Milky", "Im", "Md", "Im Md",
#         "Itemserial", "Sp", "Price", "Cts",
#         "Lab2", "Cert2", "Price/Cts1", "Price/Cts2",
#         "Price B", "6", "Im Md Itemserial Sp Price Cts"
#     ]

#     return df.drop(columns=[col for col in columns_to_remove if col in df.columns], errors='ignore')


# # ---------------- STEP 2 ----------------
# def filter_lab(df):
#     df["Lab"] = df["Lab"].astype(str).str.strip().str.upper()
#     return df[df["Lab"].isin(["GIA", "IGI", "GCAL"])]


# # ---------------- STEP 3 ----------------
# def fill_quality(df):
#     df["Quality"] = df["Quality"].fillna("").astype(str).str.strip()
#     df["Rapnet Note"] = df["Rapnet Note"].fillna("").astype(str).str.upper()

#     rapnet_map = df.set_index("Lot #")["Rapnet Note"].to_dict()

#     def update_quality(row):
#         if row["Quality"] == "":
#             rap_val = rapnet_map.get(row["Lot #"], "")
#             if "CVD" in rap_val:
#                 return "CVD"
#             elif "HPHT" in rap_val:
#                 return "HPHT"
#         return row["Quality"]

#     df["Quality"] = df.apply(update_quality, axis=1)

#     # Remove Rapnet Note
#     df = df.drop(columns=["Rapnet Note"], errors="ignore")

#     return df


# # ---------------- STEP 4 (VLOOKUP) ----------------
# def apply_vlookup(main_df, lab_file):
#     # Read Lab Grown file correctly (header row = 3)
#     lab_df = pd.read_excel(lab_file, header=2)

#     # Clean column names
#     main_df.columns = main_df.columns.str.strip()
#     lab_df.columns = lab_df.columns.str.strip()

#     # 🔍 Debug (run once if needed)
#     # st.write(lab_df.columns.tolist())

#     # Auto-detect columns
#     stock_col = [col for col in lab_df.columns if "stock" in col.lower()][0]
#     age_col = [col for col in lab_df.columns if "old" in col.lower()][0]

#     # Rename to standard
#     lab_df = lab_df.rename(columns={
#         stock_col: "Stock #",
#         age_col: "How old stone in stock"
#     })

#     # Keep only required columns
#     lab_df = lab_df[["Stock #", "How old stone in stock"]]

#     # Ensure main file has Stock #
#     if "Stock #" not in main_df.columns:
#         raise ValueError("❌ 'Stock #' column not found in Main file")

#     # Merge (VLOOKUP)
#     merged_df = pd.merge(main_df, lab_df, on="Stock #", how="left")

#     # Insert before Price / Cts
#     if "Price / Cts" in merged_df.columns:
#         cols = list(merged_df.columns)

#         age = cols.pop(cols.index("How old stone in stock"))
#         stock = cols.pop(cols.index("Stock #"))

#         idx = cols.index("Price / Cts")

#         cols.insert(idx, age)
#         cols.insert(idx, stock)

#         merged_df = merged_df[cols]

#     return merged_df


# # ---------------- MAIN ----------------
# if main_file and labgrown_file:

#     main_df = pd.read_excel(main_file)
#     lab_df = pd.read_excel(labgrown_file)

#     st.subheader("📊 Original Main Data")
#     st.dataframe(main_df.head())

#     if st.button("🚀 Process All Steps"):
#         try:
#             # Step 1 → 3
#             main_df = remove_unwanted_columns(main_df)
#             main_df = filter_lab(main_df)
#             main_df = fill_quality(main_df)

#             # Step 4
#             final_df = apply_vlookup(main_df, lab_df)

#             st.subheader("✅ Final Data After VLOOKUP")
#             st.dataframe(final_df.head())

#             # Download
#             output = io.BytesIO()
#             with pd.ExcelWriter(output, engine='openpyxl') as writer:
#                 final_df.to_excel(writer, index=False)

#             st.download_button(
#                 "📥 Download Final File",
#                 data=output.getvalue(),
#                 file_name="final_output.xlsx",
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#             )

#         except Exception as e:
#             st.error(f"❌ Error: {e}")



# Final Wokring step 4 24-04-26



# import streamlit as st
# import pandas as pd
# import io

# st.set_page_config(page_title="💎 Diamond Automation", layout="wide")

# st.title("💎 Diamond Automation (Step 1 → 4)")

# # Upload files
# main_file = st.file_uploader("Upload Main Date File", type=["xlsx"])
# labgrown_file = st.file_uploader("Upload Lab Grown File", type=["xlsx"])


# # ---------------- STEP 1 ----------------
# def remove_unwanted_columns(df):
#     df.columns = df.columns.str.strip()

#     columns_to_remove = [
#         "Polish", "Sym.", "Flu. Int.", "Tab %", "Dep %", "Cut", "Origin",
#         "List Price", "% Off",
#         "Price A%", "Price A", "Price B%", "Price B", "%RP/Cost",
#         "Rect Cost", "Other Cost", "P&L", "P&&L", "S. Qlty.",
#         "General Note", "Private Note",
#         "CN", "SN", "CW", "SW", "Milky", "Im", "Md", "Im Md",
#         "Itemserial", "Sp", "Price", "Cts",
#         "Lab2", "Cert2", "Price/Cts1", "Price/Cts2",
#         "Price B", "6", "Im Md Itemserial Sp Price Cts"
#     ]

#     return df.drop(columns=[col for col in columns_to_remove if col in df.columns], errors='ignore')


# # ---------------- STEP 2 ----------------
# def filter_lab(df):
#     df["Lab"] = df["Lab"].astype(str).str.strip().str.upper()
#     return df[df["Lab"].isin(["GIA", "IGI", "GCAL"])]


# # ---------------- STEP 3 ----------------
# def fill_quality(df):
#     df["Quality"] = df["Quality"].fillna("").astype(str).str.strip()
#     df["Rapnet Note"] = df["Rapnet Note"].fillna("").astype(str).str.upper()

#     rapnet_map = df.set_index("Lot #")["Rapnet Note"].to_dict()

#     def update_quality(row):
#         if row["Quality"] == "":
#             rap_val = rapnet_map.get(row["Lot #"], "")
#             if "CVD" in rap_val:
#                 return "CVD"
#             elif "HPHT" in rap_val:
#                 return "HPHT"
#         return row["Quality"]

#     df["Quality"] = df.apply(update_quality, axis=1)

#     # Remove Rapnet Note
#     df = df.drop(columns=["Rapnet Note"], errors="ignore")

#     return df


# # ---------------- STEP 4 ----------------
# def apply_vlookup(main_df, lab_file):
#     # Read Lab file (correct header)
#     lab_df = pd.read_excel(lab_file, header=2)

#     # Clean column names
#     main_df.columns = main_df.columns.str.strip()
#     lab_df.columns = lab_df.columns.str.strip()

#     # 🔍 Auto-detect columns
#     stock_col = [col for col in lab_df.columns if "stock" in col.lower()][0]
#     age_col = [col for col in lab_df.columns if "old" in col.lower()][0]

#     # ✅ Rename columns
#     lab_df = lab_df.rename(columns={
#         stock_col: "Lot #",
#         age_col: "No. Of Days"   # 🔥 Changed here
#     })

#     # Keep required columns
#     lab_df = lab_df[["Lot #", "No. Of Days"]]

#     # Match datatype
#     main_df["Lot #"] = main_df["Lot #"].astype(str).str.strip()
#     lab_df["Lot #"] = lab_df["Lot #"].astype(str).str.strip()

#     # ✅ Merge
#     merged_df = pd.merge(main_df, lab_df, on="Lot #", how="left")

#     # Insert before Price / Cts
#     if "Price / Cts" in merged_df.columns:
#         cols = list(merged_df.columns)

#         new_col = cols.pop(cols.index("No. Of Days"))
#         idx = cols.index("Price / Cts")

#         cols.insert(idx, new_col)

#         merged_df = merged_df[cols]

#     return merged_df


# # ---------------- MAIN ----------------
# if main_file and labgrown_file:

#     main_df = pd.read_excel(main_file)

#     st.subheader("📊 Original Main Data")
#     st.dataframe(main_df.head())

#     if st.button("🚀 Process All Steps"):
#         try:
#             # Step 1 → 3
#             main_df = remove_unwanted_columns(main_df)
#             main_df = filter_lab(main_df)
#             main_df = fill_quality(main_df)

#             # Step 4
#             final_df = apply_vlookup(main_df, labgrown_file)

#             st.subheader("✅ Final Processed Data")
#             st.dataframe(final_df.head())

#             # Download
#             output = io.BytesIO()
#             with pd.ExcelWriter(output, engine='openpyxl') as writer:
#                 final_df.to_excel(writer, index=False)

#             st.download_button(
#                 "📥 Download Final File",
#                 data=output.getvalue(),
#                 file_name="final_output.xlsx",
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#             )

#         except Exception as e:
#             st.error(f"❌ Error: {e}")



# import streamlit as st
# import pandas as pd
# import io

# st.set_page_config(page_title="💎 Diamond Automation", layout="wide")

# st.title("💎 Diamond Automation (Step 1 → 5)")

# # Upload files
# main_file = st.file_uploader("Upload Main Date File", type=["xlsx"])
# labgrown_file = st.file_uploader("Upload Lab Grown File", type=["xlsx"])
# pending_file = st.file_uploader("Upload Pending Video File", type=["xlsx"])


# # ---------------- STEP 1 ----------------
# def remove_unwanted_columns(df):
#     df.columns = df.columns.str.strip()

#     columns_to_remove = [
#         "Polish", "Sym.", "Flu. Int.", "Tab %", "Dep %", "Cut", "Origin",
#         "List Price", "% Off",
#         "Price A%", "Price A", "Price B%", "Price B", "%RP/Cost",
#         "Rect Cost", "Other Cost", "P&L", "P&&L", "S. Qlty.",
#         "General Note", "Private Note",
#         "CN", "SN", "CW", "SW", "Milky", "Im", "Md", "Im Md",
#         "Itemserial", "Sp", "Price", "Cts",
#         "Lab2", "Cert2", "Price/Cts1", "Price/Cts2",
#         "Price B", "6", "Im Md Itemserial Sp Price Cts"
#     ]

#     return df.drop(columns=[col for col in columns_to_remove if col in df.columns], errors='ignore')


# # ---------------- STEP 2 ----------------
# def filter_lab(df):
#     df["Lab"] = df["Lab"].astype(str).str.strip().str.upper()
#     return df[df["Lab"].isin(["GIA", "IGI", "GCAL"])]


# # ---------------- STEP 3 ----------------
# def fill_quality(df):
#     df["Quality"] = df["Quality"].fillna("").astype(str).str.strip()
#     df["Rapnet Note"] = df["Rapnet Note"].fillna("").astype(str).str.upper()

#     rapnet_map = df.set_index("Lot #")["Rapnet Note"].to_dict()

#     def update_quality(row):
#         if row["Quality"] == "":
#             rap_val = rapnet_map.get(row["Lot #"], "")
#             if "CVD" in rap_val:
#                 return "CVD"
#             elif "HPHT" in rap_val:
#                 return "HPHT"
#         return row["Quality"]

#     df["Quality"] = df.apply(update_quality, axis=1)

#     return df.drop(columns=["Rapnet Note"], errors="ignore")


# # ---------------- STEP 4 ----------------
# def apply_vlookup_lab(main_df, lab_file):
#     lab_df = pd.read_excel(lab_file, header=2)

#     main_df.columns = main_df.columns.str.strip()
#     lab_df.columns = lab_df.columns.str.strip()

#     stock_col = [col for col in lab_df.columns if "stock" in col.lower()][0]
#     age_col = [col for col in lab_df.columns if "old" in col.lower()][0]

#     lab_df = lab_df.rename(columns={
#         stock_col: "Lot #",
#         age_col: "No. Of Days"
#     })

#     lab_df = lab_df[["Lot #", "No. Of Days"]]

#     main_df["Lot #"] = main_df["Lot #"].astype(str).str.strip()
#     lab_df["Lot #"] = lab_df["Lot #"].astype(str).str.strip()

#     merged_df = pd.merge(main_df, lab_df, on="Lot #", how="left")

#     if "Price / Cts" in merged_df.columns:
#         cols = list(merged_df.columns)
#         new_col = cols.pop(cols.index("No. Of Days"))
#         idx = cols.index("Price / Cts")
#         cols.insert(idx, new_col)
#         merged_df = merged_df[cols]

#     return merged_df


# # ---------------- STEP 5 ----------------
# def apply_vlookup_pending(main_df, pending_file):
#     pending_df = pd.read_excel(pending_file)

#     main_df.columns = main_df.columns.str.strip()
#     pending_df.columns = pending_df.columns.str.strip()

#     # Keep only required columns
#     pending_df = pending_df[["Lot #", "Status", "Customer"]]

#     # Match datatype
#     main_df["Lot #"] = main_df["Lot #"].astype(str).str.strip()
#     pending_df["Lot #"] = pending_df["Lot #"].astype(str).str.strip()

#     # Merge
#     merged_df = pd.merge(main_df, pending_df, on="Lot #", how="left")

#     # Insert Status & Customer between Lot # and Shape
#     cols = list(merged_df.columns)

#     lot_index = cols.index("Lot #")

#     # Remove and reinsert
#     status_col = cols.pop(cols.index("Status"))
#     customer_col = cols.pop(cols.index("Customer"))

#     cols.insert(lot_index + 1, status_col)
#     cols.insert(lot_index + 2, customer_col)

#     merged_df = merged_df[cols]

#     return merged_df


# # ---------------- MAIN ----------------
# if main_file and labgrown_file and pending_file:

#     main_df = pd.read_excel(main_file)

#     st.subheader("📊 Original Main Data")
#     st.dataframe(main_df.head())

#     if st.button("🚀 Process All Steps"):
#         try:
#             # Step 1 → 3
#             main_df = remove_unwanted_columns(main_df)
#             main_df = filter_lab(main_df)
#             main_df = fill_quality(main_df)

#             # Step 4
#             main_df = apply_vlookup_lab(main_df, labgrown_file)

#             # Step 5
#             final_df = apply_vlookup_pending(main_df, pending_file)

#             st.subheader("✅ Final Processed Data")
#             st.dataframe(final_df.head())

#             # Download
#             output = io.BytesIO()
#             with pd.ExcelWriter(output, engine='openpyxl') as writer:
#                 final_df.to_excel(writer, index=False)

#             st.download_button(
#                 "📥 Download Final File",
#                 data=output.getvalue(),
#                 file_name="final_output.xlsx",
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#             )

#         except Exception as e:
#             st.error(f"❌ Error: {e}")


# import streamlit as st
# import pandas as pd
# import io

# st.set_page_config(page_title="💎 Diamond Automation", layout="wide")

# st.title("💎 Diamond Automation (Step 1 → 6)")

# # Upload files
# main_file = st.file_uploader("Upload Main Date File", type=["xlsx"])
# labgrown_file = st.file_uploader("Upload Lab Grown File", type=["xlsx"])
# pending_file = st.file_uploader("Upload Pending Video File", type=["xlsx"])


# # ---------------- STEP 1 ----------------
# def remove_unwanted_columns(df):
#     df.columns = df.columns.str.strip()

#     columns_to_remove = [
#         "Polish", "Sym.", "Flu. Int.", "Tab %", "Dep %", "Cut", "Origin",
#         "List Price", "% Off",
#         "Price A%", "Price A", "Price B%", "Price B", "%RP/Cost",
#         "Rect Cost", "Other Cost", "P&L", "P&&L", "S. Qlty.",
#         "General Note", "Private Note",
#         "CN", "SN", "CW", "SW", "Milky", "Im", "Md", "Im Md",
#         "Itemserial", "Sp", "Price", "Cts",
#         "Lab2", "Cert2", "Price/Cts1", "Price/Cts2",
#         "Price B", "6", "Im Md Itemserial Sp Price Cts"
#     ]

#     return df.drop(columns=[col for col in columns_to_remove if col in df.columns], errors='ignore')


# # ---------------- STEP 2 ----------------
# def filter_lab(df):
#     df["Lab"] = df["Lab"].astype(str).str.strip().str.upper()
#     return df[df["Lab"].isin(["GIA", "IGI", "GCAL"])]


# # ---------------- STEP 3 ----------------
# def fill_quality(df):
#     df["Quality"] = df["Quality"].fillna("").astype(str).str.strip()
#     df["Rapnet Note"] = df["Rapnet Note"].fillna("").astype(str).str.upper()

#     rapnet_map = df.set_index("Lot #")["Rapnet Note"].to_dict()

#     def update_quality(row):
#         if row["Quality"] == "":
#             rap_val = rapnet_map.get(row["Lot #"], "")
#             if "CVD" in rap_val:
#                 return "CVD"
#             elif "HPHT" in rap_val:
#                 return "HPHT"
#         return row["Quality"]

#     df["Quality"] = df.apply(update_quality, axis=1)

#     return df.drop(columns=["Rapnet Note"], errors="ignore")


# # ---------------- STEP 4 ----------------
# def apply_vlookup_lab(main_df, lab_file):
#     lab_df = pd.read_excel(lab_file, header=2)

#     main_df.columns = main_df.columns.str.strip()
#     lab_df.columns = lab_df.columns.str.strip()

#     stock_col = [col for col in lab_df.columns if "stock" in col.lower()][0]
#     age_col = [col for col in lab_df.columns if "old" in col.lower()][0]

#     lab_df = lab_df.rename(columns={
#         stock_col: "Lot #",
#         age_col: "No. Of Days"
#     })

#     lab_df = lab_df[["Lot #", "No. Of Days"]]

#     main_df["Lot #"] = main_df["Lot #"].astype(str).str.strip()
#     lab_df["Lot #"] = lab_df["Lot #"].astype(str).str.strip()

#     merged_df = pd.merge(main_df, lab_df, on="Lot #", how="left")

#     if "Price / Cts" in merged_df.columns:
#         cols = list(merged_df.columns)
#         new_col = cols.pop(cols.index("No. Of Days"))
#         idx = cols.index("Price / Cts")
#         cols.insert(idx, new_col)
#         merged_df = merged_df[cols]

#     return merged_df


# # ---------------- STEP 5 ----------------
# def apply_vlookup_pending(main_df, pending_file):
#     pending_df = pd.read_excel(pending_file)

#     main_df.columns = main_df.columns.str.strip()
#     pending_df.columns = pending_df.columns.str.strip()

#     pending_df = pending_df[["Lot #", "Status", "Customer"]]

#     main_df["Lot #"] = main_df["Lot #"].astype(str).str.strip()
#     pending_df["Lot #"] = pending_df["Lot #"].astype(str).str.strip()

#     merged_df = pd.merge(main_df, pending_df, on="Lot #", how="left")

#     # Insert between Lot # and Shape
#     cols = list(merged_df.columns)
#     lot_index = cols.index("Lot #")

#     status_col = cols.pop(cols.index("Status"))
#     customer_col = cols.pop(cols.index("Customer"))

#     cols.insert(lot_index + 1, status_col)
#     cols.insert(lot_index + 2, customer_col)

#     merged_df = merged_df[cols]

#     return merged_df


# # ---------------- STEP 6 ----------------
# def update_status_and_cleanup(df):
#     df["Customer"] = df["Customer"].fillna("").str.upper()
#     df["Status"] = df["Status"].fillna("").str.strip()

#     # Condition
#     mask = df["Customer"].isin([
#         "GOODS IN TRANSIT",
#         "GOODS IN TRANSIT FROM OVERSEAS"
#     ]) & (df["Status"].str.upper() == "ONMEMO")

#     df.loc[mask, "Status"] = "Inhand"

#     # Remove Customer column
#     df = df.drop(columns=["Customer"], errors="ignore")

#     return df


# # ---------------- MAIN ----------------
# if main_file and labgrown_file and pending_file:

#     main_df = pd.read_excel(main_file)

#     st.subheader("📊 Original Main Data")
#     st.dataframe(main_df.head())

#     if st.button("🚀 Process All Steps"):
#         try:
#             # Step 1 → 3
#             main_df = remove_unwanted_columns(main_df)
#             main_df = filter_lab(main_df)
#             main_df = fill_quality(main_df)

#             # Step 4
#             main_df = apply_vlookup_lab(main_df, labgrown_file)

#             # Step 5
#             main_df = apply_vlookup_pending(main_df, pending_file)

#             # Step 6
#             final_df = update_status_and_cleanup(main_df)

#             st.subheader("✅ Final Processed Data")
#             st.dataframe(final_df.head())

#             # Download
#             output = io.BytesIO()
#             with pd.ExcelWriter(output, engine='openpyxl') as writer:
#                 final_df.to_excel(writer, index=False)

#             st.download_button(
#                 "📥 Download Final File",
#                 data=output.getvalue(),
#                 file_name="final_output.xlsx",
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#             )

#         except Exception as e:
#             st.error(f"❌ Error: {e}")

# # 
# import streamlit as st
# import pandas as pd
# import io

# st.set_page_config(page_title="💎 Diamond Automation", layout="wide")

# st.title("💎 Diamond Automation (Step 1 → 7)")

# # Upload files
# main_file = st.file_uploader("Upload Main Date File", type=["xlsx"])
# labgrown_file = st.file_uploader("Upload Lab Grown File", type=["xlsx"])
# pending_file = st.file_uploader("Upload Pending Video File", type=["xlsx"])


# # ---------------- STEP 1 ----------------
# def remove_unwanted_columns(df):
#     df.columns = df.columns.str.strip()

#     columns_to_remove = [
#         "Polish", "Sym.", "Flu. Int.", "Tab %", "Dep %", "Cut", "Origin",
#         "List Price", "% Off",
#         "Price A%", "Price A", "Price B%", "Price B", "%RP/Cost",
#         "Rect Cost", "Other Cost", "P&L", "P&&L", "S. Qlty.",
#         "General Note", "Private Note",
#         "CN", "SN", "CW", "SW", "Milky", "Im", "Md", "Im Md",
#         "Itemserial", "Sp", "Price", "Cts",
#         "Lab2", "Cert2", "Price/Cts1", "Price/Cts2",
#         "Price B", "6", "Im Md Itemserial Sp Price Cts"
#     ]

#     return df.drop(columns=[col for col in columns_to_remove if col in df.columns], errors='ignore')


# # ---------------- STEP 2 ----------------
# def filter_lab(df):
#     df["Lab"] = df["Lab"].astype(str).str.strip().str.upper()
#     return df[df["Lab"].isin(["GIA", "IGI", "GCAL"])]


# # ---------------- STEP 3 ----------------
# def fill_quality(df):
#     df["Quality"] = df["Quality"].fillna("").astype(str).str.strip()
#     df["Rapnet Note"] = df["Rapnet Note"].fillna("").astype(str).str.upper()

#     rapnet_map = df.set_index("Lot #")["Rapnet Note"].to_dict()

#     def update_quality(row):
#         if row["Quality"] == "":
#             rap_val = rapnet_map.get(row["Lot #"], "")
#             if "CVD" in rap_val:
#                 return "CVD"
#             elif "HPHT" in rap_val:
#                 return "HPHT"
#         return row["Quality"]

#     df["Quality"] = df.apply(update_quality, axis=1)

#     return df.drop(columns=["Rapnet Note"], errors="ignore")


# # ---------------- STEP 4 ----------------
# def apply_vlookup_lab(main_df, lab_file):
#     lab_df = pd.read_excel(lab_file, header=2)

#     main_df.columns = main_df.columns.str.strip()
#     lab_df.columns = lab_df.columns.str.strip()

#     stock_col = [col for col in lab_df.columns if "stock" in col.lower()][0]
#     age_col = [col for col in lab_df.columns if "old" in col.lower()][0]

#     lab_df = lab_df.rename(columns={
#         stock_col: "Lot #",
#         age_col: "No. Of Days"
#     })

#     lab_df = lab_df[["Lot #", "No. Of Days"]]

#     main_df["Lot #"] = main_df["Lot #"].astype(str).str.strip()
#     lab_df["Lot #"] = lab_df["Lot #"].astype(str).str.strip()

#     merged_df = pd.merge(main_df, lab_df, on="Lot #", how="left")

#     if "Price / Cts" in merged_df.columns:
#         cols = list(merged_df.columns)
#         new_col = cols.pop(cols.index("No. Of Days"))
#         idx = cols.index("Price / Cts")
#         cols.insert(idx, new_col)
#         merged_df = merged_df[cols]

#     return merged_df


# # ---------------- STEP 5 ----------------
# def apply_vlookup_pending(main_df, pending_file):
#     pending_df = pd.read_excel(pending_file)

#     main_df.columns = main_df.columns.str.strip()
#     pending_df.columns = pending_df.columns.str.strip()

#     pending_df = pending_df[["Lot #", "Status", "Customer"]]

#     main_df["Lot #"] = main_df["Lot #"].astype(str).str.strip()
#     pending_df["Lot #"] = pending_df["Lot #"].astype(str).str.strip()

#     merged_df = pd.merge(main_df, pending_df, on="Lot #", how="left")

#     cols = list(merged_df.columns)
#     lot_index = cols.index("Lot #")

#     status_col = cols.pop(cols.index("Status"))
#     customer_col = cols.pop(cols.index("Customer"))

#     cols.insert(lot_index + 1, status_col)
#     cols.insert(lot_index + 2, customer_col)

#     return merged_df[cols]


# # ---------------- STEP 6 ----------------
# def update_status_and_cleanup(df):
#     df["Customer"] = df["Customer"].fillna("").str.upper()
#     df["Status"] = df["Status"].fillna("").str.strip()

#     mask = df["Customer"].isin([
#         "GOODS IN TRANSIT",
#         "GOODS IN TRANSIT FROM OVERSEAS"
#     ]) & (df["Status"].str.upper() == "ONMEMO")

#     df.loc[mask, "Status"] = "Inhand"

#     return df.drop(columns=["Customer"], errors="ignore")


# # ---------------- STEP 7 ----------------
# def split_by_person(df):
#     df["Shape"] = df["Shape"].astype(str).str.upper().str.strip()

#     # 🔥 Handle RBC as ROUND
#     df["Shape"] = df["Shape"].replace({
#         "RBC": "ROUND"
#     })

#     mapping = {
#         "Love": ["ASSCHER", "PRINCESS", "ROUND"],
#         "Milan": ["PEAR", "RADIANT"],
#         "Gautam": ["EMERALD", "OVAL"],
#         "Girl": ["CUSHION MODIFIED", "BRILLIANT", "HEART"]
#     }

#     files = {}

#     for person, shapes in mapping.items():
#         person_df = df[df["Shape"].isin(shapes)]

#         if not person_df.empty:
#             output = io.BytesIO()
#             with pd.ExcelWriter(output, engine='openpyxl') as writer:
#                 person_df.to_excel(writer, index=False)

#             files[person] = output.getvalue()

#     return files


# # ---------------- MAIN ----------------
# if main_file and labgrown_file and pending_file:

#     main_df = pd.read_excel(main_file)

#     st.subheader("📊 Original Main Data")
#     st.dataframe(main_df.head())

#     if st.button("🚀 Process All Steps"):
#         try:
#             # Step 1 → 3
#             main_df = remove_unwanted_columns(main_df)
#             main_df = filter_lab(main_df)
#             main_df = fill_quality(main_df)

#             # Step 4
#             main_df = apply_vlookup_lab(main_df, labgrown_file)

#             # Step 5
#             main_df = apply_vlookup_pending(main_df, pending_file)

#             # Step 6
#             final_df = update_status_and_cleanup(main_df)

#             st.subheader("✅ Final Processed Data")
#             st.dataframe(final_df.head())

#             # Download final file
#             output = io.BytesIO()
#             with pd.ExcelWriter(output, engine='openpyxl') as writer:
#                 final_df.to_excel(writer, index=False)

#             st.download_button(
#                 "📥 Download Final File",
#                 data=output.getvalue(),
#                 file_name="final_output.xlsx",
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#             )

#             # Step 7 split
#             split_files = split_by_person(final_df)

#             st.subheader("📂 Download Person-wise Files")

#             for person, file_data in split_files.items():
#                 st.download_button(
#                     label=f"📥 Download {person} File",
#                     data=file_data,
#                     file_name=f"{person}_diamonds.xlsx",
#                     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#                 )

#         except Exception as e:
#             st.error(f"❌ Error: {e}")


import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="💎 Diamond Automation", layout="wide")

st.title("💎 Diamond Automation (Step 1 → 6)")

# Upload files
main_file = st.file_uploader("Upload Main Date File", type=["xlsx"])
labgrown_file = st.file_uploader("Upload Lab Grown File", type=["xlsx"])
pending_file = st.file_uploader("Upload Pending Video File", type=["xlsx"])


# ---------------- STEP 1 ----------------
def remove_unwanted_columns(df):
    df.columns = df.columns.str.strip()

    columns_to_remove = [
        "Polish", "Sym.", "Flu. Int.", "Tab %", "Dep %", "Cut", "Origin",
        "List Price", "% Off",
        "Price A%", "Price A", "Price B%", "Price B", "%RP/Cost",
        "Rect Cost", "Other Cost", "P&L", "P&&L", "S. Qlty.",
        "General Note", "Private Note",
        "CN", "SN", "CW", "SW", "Milky", "Im", "Md", "Im Md",
        "Itemserial", "Sp", "Price", "Cts",
        "Lab2", "Cert2", "Price/Cts1", "Price/Cts2",
        "Price B", "6", "Im Md Itemserial Sp Price Cts"
    ]

    return df.drop(columns=[col for col in columns_to_remove if col in df.columns], errors='ignore')


# ---------------- STEP 2 ----------------
def filter_lab(df):
    df["Lab"] = df["Lab"].astype(str).str.strip().str.upper()
    return df[df["Lab"].isin(["GIA", "IGI", "GCAL"])]


# ---------------- STEP 3 ----------------
def fill_quality(df):
    df["Quality"] = df["Quality"].fillna("").astype(str).str.strip()
    df["Rapnet Note"] = df["Rapnet Note"].fillna("").astype(str).str.upper()

    rapnet_map = df.set_index("Lot #")["Rapnet Note"].to_dict()

    def update_quality(row):
        if row["Quality"] == "":
            rap_val = rapnet_map.get(row["Lot #"], "")
            if "CVD" in rap_val:
                return "CVD"
            elif "HPHT" in rap_val:
                return "HPHT"
        return row["Quality"]

    df["Quality"] = df.apply(update_quality, axis=1)

    return df.drop(columns=["Rapnet Note"], errors="ignore")


# ---------------- STEP 4 ----------------
def apply_vlookup_lab(main_df, lab_file):
    lab_df = pd.read_excel(lab_file, header=2)

    main_df.columns = main_df.columns.str.strip()
    lab_df.columns = lab_df.columns.str.strip()

    stock_col = [col for col in lab_df.columns if "stock" in col.lower()][0]
    age_col = [col for col in lab_df.columns if "old" in col.lower()][0]

    lab_df = lab_df.rename(columns={
        stock_col: "Lot #",
        age_col: "No. Of Days"
    })

    lab_df = lab_df[["Lot #", "No. Of Days"]]

    main_df["Lot #"] = main_df["Lot #"].astype(str).str.strip()
    lab_df["Lot #"] = lab_df["Lot #"].astype(str).str.strip()

    merged_df = pd.merge(main_df, lab_df, on="Lot #", how="left")

    if "Price / Cts" in merged_df.columns:
        cols = list(merged_df.columns)
        new_col = cols.pop(cols.index("No. Of Days"))
        idx = cols.index("Price / Cts")
        cols.insert(idx, new_col)
        merged_df = merged_df[cols]

    return merged_df


# ---------------- STEP 5 ----------------
def apply_vlookup_pending(main_df, pending_file):
    pending_df = pd.read_excel(pending_file)

    main_df.columns = main_df.columns.str.strip()
    pending_df.columns = pending_df.columns.str.strip()

    pending_df = pending_df[["Lot #", "Status", "Customer"]]

    main_df["Lot #"] = main_df["Lot #"].astype(str).str.strip()
    pending_df["Lot #"] = pending_df["Lot #"].astype(str).str.strip()

    merged_df = pd.merge(main_df, pending_df, on="Lot #", how="left")

    # Insert between Lot # and Shape
    cols = list(merged_df.columns)
    lot_index = cols.index("Lot #")

    status_col = cols.pop(cols.index("Status"))
    customer_col = cols.pop(cols.index("Customer"))

    cols.insert(lot_index + 1, status_col)
    cols.insert(lot_index + 2, customer_col)

    merged_df = merged_df[cols]

    return merged_df


# ---------------- STEP 6 ----------------
def update_status_and_cleanup(df):
    df["Customer"] = df["Customer"].fillna("").str.upper()
    df["Status"] = df["Status"].fillna("").str.strip()

    # Condition
    mask = df["Customer"].isin([
        "GOODS IN TRANSIT",
        "GOODS IN TRANSIT FROM OVERSEAS"
    ]) & (df["Status"].str.upper() == "ONMEMO")

    df.loc[mask, "Status"] = "Inhand"

    # Remove Customer column
    df = df.drop(columns=["Customer"], errors="ignore")

    return df


# ---------------- MAIN ----------------
if main_file and labgrown_file and pending_file:

    main_df = pd.read_excel(main_file)

    st.subheader("📊 Original Main Data")
    st.dataframe(main_df.head())

    if st.button("🚀 Process All Steps"):
        try:
            # Step 1 → 3
            main_df = remove_unwanted_columns(main_df)
            main_df = filter_lab(main_df)
            main_df = fill_quality(main_df)

            # Step 4
            main_df = apply_vlookup_lab(main_df, labgrown_file)

            # Step 5
            main_df = apply_vlookup_pending(main_df, pending_file)

            # Step 6
            final_df = update_status_and_cleanup(main_df)

            st.subheader("✅ Final Processed Data")
            st.dataframe(final_df.head())

            # Download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False)

            st.download_button(
                "📥 Download Final File",
                data=output.getvalue(),
                file_name="final_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Error: {e}")

# LAST UPADTE


