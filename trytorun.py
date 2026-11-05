import io
import os
import pandas as pd
import streamlit as st

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="SellEthics Analytics Calculator",
    page_icon="📊",
    layout="wide",
)

# --- CUSTOM CSS STYLING (Dark Theme Matching) ---
st.markdown(
    """
    <style>
    .stApp {
        background-color: #181825;
        color: #CDD6F4;
    }
    sidebar .stSidebar {
        background-color: #1E1E2E;
    }
    .metric-card {
        background-color: #313244;
        padding: 15px;
        border-radius: 8px;
        text-align: center;
        font-weight: bold;
        font-size: 18px;
        color: #89B4FA;
    }
    </style>
""",
    unsafe_allow_html=True,
)


# --- HELPER FUNCTIONS ---
def clean_numeric_input(val_str):
  """Sanitizes strings containing $, commas, spaces, or % symbols."""
  if not val_str:
    return 0.0
  cleaned = (
      str(val_str).replace("$", "").replace(",", "").replace("%", "").strip()
  )
  try:
    return float(cleaned)
  except ValueError:
    return 0.0


def get_fsi_color(fsi_val):
  if fsi_val <= 75:
    return "#14532d", "#86efac"
  elif fsi_val <= 99:
    return "#166534", "#bbf7d0"
  elif fsi_val <= 105:
    return "#713f12", "#fef08a"
  elif fsi_val <= 115:
    return "#854d0e", "#fde047"
  elif fsi_val <= 130:
    return "#7c2d12", "#fdba74"
  elif fsi_val <= 150:
    return "#7f1d1d", "#fca5a5"
  else:
    return "#ef4444", "#ffffff"


def determine_mechanic(promo_aup, base_aup):
  raw_discount = base_aup - promo_aup
  rounded_tpr = round(raw_discount * 20) / 20.0
  discount_pct = (1 - (promo_aup / base_aup)) * 100 if base_aup > 0 else 0

  if 45.0 <= discount_pct <= 55.0:
    return "BOGO"
  elif 1.85 <= promo_aup <= 2.15 and base_aup >= 2.40:
    return "2 for $4"
  elif 2.35 <= promo_aup <= 2.65 and base_aup >= 2.90:
    return "2 for $5"
  elif 2.80 <= promo_aup <= 3.20 and base_aup >= 3.40:
    return "2 for $6"
  elif 3.30 <= promo_aup <= 3.70 and base_aup >= 3.90:
    return "2 for $7"
  elif 4.75 <= promo_aup <= 5.25 and base_aup >= 5.60:
    return "2 for $10"
  elif discount_pct > 5:
    return f"~${rounded_tpr:.2f} TPR"
  else:
    return "Display / Feature Only"


def format_date_str(val):
  try:
    dt = pd.to_datetime(val)
    return dt.strftime("%m/%d/%Y")
  except Exception:
    s = str(val).split(" ")[0]
    return s.replace("-", "/")


def check_sheet_contains_headers(raw_df):
  try:
    block_str = " ".join([
        str(val).lower().replace("\n", " ")
        for row in raw_df.values
        for val in row
        if pd.notna(val)
    ])
    return "promo" in block_str and "incr" in block_str and "base" in block_str
  except Exception:
    return False


def find_valid_sheet_and_names(file_bytes):
  try:
    xls = pd.ExcelFile(file_bytes)
    sheet_names = xls.sheet_names
    for s_name in sheet_names:
      try:
        raw_df = pd.read_excel(xls, sheet_name=s_name, header=None, nrows=25)
        if check_sheet_contains_headers(raw_df):
          return sheet_names, str(s_name)
      except Exception:
        continue
    return sheet_names, str(sheet_names[0])
  except Exception:
    return ["Sheet1"], "Sheet1"


def find_header_row_and_df(xls, sheet_name):
  try:
    raw_df = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=25)
    header_row = 0
    for idx, row in raw_df.iterrows():
      row_str = " ".join([
          str(val).lower().replace("\n", " ") for val in row.values if pd.notna(val)
      ])
      if "unit" in row_str or "base" in row_str or "incr" in row_str:
        header_row = idx
        break
    return pd.read_excel(xls, sheet_name=sheet_name, header=header_row)
  except Exception:
    return pd.read_excel(xls, sheet_name=sheet_name, header=0)


# --- SIDEBAR NAVIGATION ---
st.sidebar.title("Analytics Calculator")
module = st.sidebar.radio(
    "Select Module",
    [
        "Home",
        "Forecasting Increased Distribution",
        "Forecasting New Item Distribution",
        "Fair Share Index",
        "Opportunity Gap",
        "Promotional Analysis",
    ],
)

weeks_map = {"L13": 13, "L26": 26, "L52": 52}

# ==========================================
# 1. HOME MODULE
# ==========================================
if module == "Home":
  st.markdown(
      "<h1 style='text-align: center; margin-top: 100px;'>Analytics"
      " Calculator</h1>",
      unsafe_allow_html=True,
  )
  st.markdown(
      "<p style='text-align: center; color: #A6ADC8;'>Select a module from the"
      " left sidebar navigation, enter your parameters, and perform analytics"
      " calculations!</p>",
      unsafe_allow_html=True,
  )

# ==========================================
# 2. FORECASTING INCREASED DISTRIBUTION
# ==========================================
elif module == "Forecasting Increased Distribution":
  st.header("Forecasting Increased Distribution")

  col1, col2 = st.columns(2)
  with col1:
    actual_period = st.selectbox(
        "Choose Actual Time Period", ["L13", "L26", "L52"], key="inc_act"
    )
  with col2:
    forecast_period = st.selectbox(
        "Choose Forecast Time Period", ["L13", "L26", "L52"], key="inc_fc"
    )

  sales_val = st.text_input(
      "Enter Actual $ Sales", placeholder="e.g. 50000", key="inc_sales"
  )

  driver_type = st.radio(
      "Choose Distribution Metric", ["%ACV Reach", "TDP"], key="inc_driver"
  )

  if driver_type == "%ACV Reach":
    metric_val = st.text_input(
        "Enter %ACV Reach Value", placeholder="e.g. 75.5", key="inc_acv"
    )
  else:
    metric_val = st.text_input(
        "Enter TDP Value", placeholder="e.g. 120", key="inc_tdp"
    )

  if st.button("Calculate", type="primary"):
    try:
      sales = clean_numeric_input(sales_val)
      m_val = clean_numeric_input(metric_val)

      if m_val == 0:
        st.error("Distribution metric value cannot be zero.")
      else:
        act_w = weeks_map[actual_period]
        fc_w = weeks_map[forecast_period]
        time_multiplier = fc_w / act_w

        scaled_actual_sales = sales * time_multiplier
        new_projected_amount = (sales / m_val) * 100 * time_multiplier
        increased_amount = new_projected_amount - scaled_actual_sales

        st.markdown("### Results")
        c1, c2 = st.columns(2)
        with c1:
          st.metric(
              label="New Projected Amount", value=f"${new_projected_amount:,.2f}"
          )
        with c2:
          st.metric(label="Increased Amount", value=f"${increased_amount:,.2f}")
    except Exception as e:
      st.error(f"Please enter valid numeric values. Error: {e}")

# ==========================================
# 3. FORECASTING NEW ITEM DISTRIBUTION
# ==========================================
elif module == "Forecasting New Item Distribution":
  st.header("Forecasting New Item Distribution")

  col1, col2 = st.columns(2)
  with col1:
    actual_period = st.selectbox(
        "Choose Actual Time Period", ["L13", "L26", "L52"], key="new_act"
    )
  with col2:
    forecast_period = st.selectbox(
        "Choose Forecast Time Period", ["L13", "L26", "L52"], key="new_fc"
    )

  sales_val = st.text_input(
      "Enter Actual $ Sales", placeholder="e.g. 25000", key="new_sales"
  )

  col3, col4 = st.columns(2)
  with col3:
    comp_acv = st.text_input("COMP $/$MM ACV", key="new_comp")
    ta_acv = st.text_input("TA Est MKT ACV", key="new_ta")
  with col4:
    tdp_val = st.text_input("TDP", key="new_tdp")
    share_val = st.text_input("Retailer's Share of Market %", key="new_share")

  if st.button("Calculate", type="primary"):
    try:
      sales = clean_numeric_input(sales_val)
      c_acv = clean_numeric_input(comp_acv)
      t_acv = clean_numeric_input(ta_acv)
      tdp = clean_numeric_input(tdp_val)
      share = clean_numeric_input(share_val)

      if tdp == 0:
        st.error("TDP cannot be zero.")
      else:
        act_w = weeks_map[actual_period]
        fc_w = weeks_map[forecast_period]
        time_multiplier = fc_w / act_w

        scaled_sales = sales * time_multiplier
        left_val = (c_acv * t_acv) * time_multiplier
        distribution_points = scaled_sales / tdp
        right_val = distribution_points * share
        final_forecast = (left_val + right_val) / 2

        st.markdown("### Results")
        st.metric(label="Forecasted $ Sales", value=f"${final_forecast:,.2f}")
    except Exception as e:
      st.error(f"Please enter valid numeric values. Error: {e}")

# ==========================================
# 4. FAIR SHARE INDEX
# ==========================================
elif module == "Fair Share Index":
  st.header("Fair Share Index")

  brand_sales = st.text_input("Brand $ Sales", key="fs_bs")
  total_sales = st.text_input("Total Category $ Sales", key="fs_ts")
  brand_tdp = st.text_input("Brand TDP", key="fs_bt")
  total_tdp = st.text_input("Total Category TDP", key="fs_tt")

  if st.button("Calculate", type="primary"):
    try:
      b_sales = clean_numeric_input(brand_sales)
      t_sales = clean_numeric_input(total_sales)
      b_tdp = clean_numeric_input(brand_tdp)
      t_tdp = clean_numeric_input(total_tdp)

      if t_sales == 0 or t_tdp == 0 or b_sales == 0:
        st.error("Sales and TDP values cannot be zero.")
      else:
        brand_dollar_share = (b_sales / t_sales) * 100
        tdp_share = (b_tdp / t_tdp) * 100
        fair_share_index = (tdp_share / brand_dollar_share) * 100

        st.markdown("### Results")
        st.metric(label="Fair Share Index", value=f"{fair_share_index:,.1f}")
    except Exception as e:
      st.error(f"Please enter valid numeric values. Error: {e}")

# ==========================================
# 5. OPPORTUNITY GAP
# ==========================================
elif module == "Opportunity Gap":
  st.header("Opportunity Gap")

  ret_acv = st.text_input(
      "Retailer Est Market ACV ($MM)", key="og_racv"
  )
  comp_acv = st.text_input(
      "Comp Market Est Market ACV ($MM)", key="og_cacv"
  )
  ret_sales = st.text_input("Retailer $ Sales", key="og_rsales")
  comp_sales = st.text_input("Comp Market $ Sales", key="og_csales")

  if st.button("Calculate", type="primary"):
    try:
      r_acv = clean_numeric_input(ret_acv)
      c_acv = clean_numeric_input(comp_acv)
      r_sales = clean_numeric_input(ret_sales)
      c_sales = clean_numeric_input(comp_sales)

      if c_acv == 0:
        st.error("Comp Market Est Market ACV cannot be zero.")
      else:
        share_of_market = round(r_acv / c_acv, 3)
        opportunity_gap = r_sales - (c_sales * share_of_market)

        st.markdown("### Results")
        st.metric(label="Opportunity Gap", value=f"${opportunity_gap:,.0f}")
    except Exception as e:
      st.error(f"Please enter valid numeric values. Error: {e}")

# ==========================================
# 6. PROMOTIONAL ANALYSIS
# ==========================================
elif module == "Promotional Analysis":
  st.header("Promotional Analysis")
  st.markdown(
      "*Note: Please ensure your uploaded file includes columns for: % Units Any"
      " Promo, Incr Units, Base Units, AUP, and Base AUP.*"
  )

  uploaded_file = st.file_uploader(
      "Upload Excel or CSV Report", type=["xlsx", "csv"]
  )
  promo_threshold = st.slider(
      "Promo Threshold",
      min_value=0.50,
      max_value=1.00,
      value=0.89,
      step=0.01,
      format="%d%%",
  )

  if uploaded_file is not None:
    try:
      file_bytes = uploaded_file.getvalue()
      sheet_name = None

      if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes))
        sheet_name = "CSV File"
      else:
        sheet_names, best_sheet = find_valid_sheet_and_names(
            io.BytesIO(file_bytes)
        )
        if len(sheet_names) > 1:
          sheet_name = st.selectbox(
              "Select Sheet", sheet_names, index=sheet_names.index(best_sheet)
          )
        else:
          sheet_name = best_sheet
        df = find_header_row_and_df(io.BytesIO(file_bytes), sheet_name)

      df.columns = df.columns.str.strip()
      first_col = df.columns[0]
      df = df[df[first_col].notna()].copy()

      col_map = {str(c).lower().replace("\n", " "): c for c in df.columns}
      promo_col_key = next(
          (
              col_map[k]
              for k in col_map
              if ("%" in k or "pct" in k or "percent" in k)
              and "promo" in k
              and "unit" in k
          ),
          None,
      )
      if not promo_col_key:
        promo_col_key = next(
            (
                col_map[k]
                for k in col_map
                if "promo" in k and "unit" in k and "any" not in k
            ),
            None,
        )

      incr_units_key = next(
          (col_map[k] for k in col_map if "incr" in k and "unit" in k), None
      )
      base_units_key = next(
          (col_map[k] for k in col_map if "base" in k and "unit" in k), None
      )
      aup_key = next(
          (
              col_map[k]
              for k in col_map
              if k == "aup"
              or ("aup" in k and "base" not in k and "ya" not in k)
          ),
          None,
      )
      base_aup_key = next(
          (
              col_map[k]
              for k in col_map
              if "base aup" in k or ("base" in k and "aup" in k)
          ),
          None,
      )

      if not promo_col_key or not incr_units_key or not base_units_key:
        st.error(
            "Could not find expected columns (% Units Any Promo, Incr Units,"
            " Base Units)."
        )
      else:
        for col in [incr_units_key, base_units_key]:
          df[col] = pd.to_numeric(
              df[col]
              .astype(str)
              .str.replace(",", "")
              .str.replace("$", "")
              .str.replace("%", "")
              .str.strip(),
              errors="coerce",
          ).fillna(0)

        df[promo_col_key] = pd.to_numeric(
            df[promo_col_key]
            .astype(str)
            .str.replace(",", "")
            .str.replace("$", "")
            .str.replace("%", "")
            .str.strip(),
            errors="coerce",
        ).fillna(0)
        if df[promo_col_key].max() > 1.0:
          df[promo_col_key] = df[promo_col_key] / 100.0
        df[promo_col_key] = df[promo_col_key].round(4)

        if aup_key:
          df[aup_key] = pd.to_numeric(
              df[aup_key]
              .astype(str)
              .str.replace("$", "")
              .str.replace(",", "")
              .str.strip(),
              errors="coerce",
          )
        if base_aup_key:
          df[base_aup_key] = pd.to_numeric(
              df[base_aup_key]
              .astype(str)
              .str.replace("$", "")
              .str.replace(",", "")
              .str.strip(),
              errors="coerce",
          )

        total_weeks = len(df)
        df["is_promo"] = df[promo_col_key] >= promo_threshold
        promo_weeks_count = df["is_promo"].sum()

        groups = []
        current_group = []
        for idx, row in df.iterrows():
          if row["is_promo"]:
            if not current_group:
              current_group.append(row)
            else:
              prev_row = current_group[-1]
              is_consecutive = idx == prev_row.name + 1
              price_similar = True
              if (
                  aup_key
                  and pd.notna(row[aup_key])
                  and pd.notna(prev_row[aup_key])
              ):
                if abs(row[aup_key] - prev_row[aup_key]) > 0.25:
                  price_similar = False
              if is_consecutive and price_similar:
                current_group.append(row)
              else:
                groups.append(pd.DataFrame(current_group))
                current_group = [row]
          else:
            if current_group:
              groups.append(pd.DataFrame(current_group))
              current_group = []
        if current_group:
          groups.append(pd.DataFrame(current_group))

        parsed_groups_records = []
        total_all_base_units = 0
        total_all_incr_units = 0

        for i, grp in enumerate(groups, 1):
          start_date = format_date_str(grp.iloc[0][first_col])
          end_date = format_date_str(grp.iloc[-1][first_col])
          date_range_str = (
              f"{start_date}"
              if len(grp) == 1
              else f"{start_date} – {end_date}"
          )

          grp_base_units = grp[base_units_key].sum()
          grp_incr_units = grp[incr_units_key].sum()
          grp_avg_incr = grp[incr_units_key].mean()
          grp_lift_pct = (
              (grp_incr_units / grp_base_units * 100)
              if grp_base_units > 0
              else 0
          )

          total_all_base_units += grp_base_units
          total_all_incr_units += grp_incr_units

          mech = "Standard TPR"
          promoted_aup_val = 0.0
          base_aup_val = 0.0
          if aup_key and base_aup_key:
            promoted_aup_val = grp[aup_key].mean()
            base_aup_val = grp[base_aup_key].mean()
            mech = determine_mechanic(promoted_aup_val, base_aup_val)

          parsed_groups_records.append({
              "Group": f"Group {i}",
              "Date Range": date_range_str,
              "Duration (Wks)": len(grp),
              "Promo Type": mech,
              "Promoted AUP": (
                  f"${promoted_aup_val:.2f}" if promoted_aup_val else "N/A"
              ),
              "Base AUP": f"${base_aup_val:.2f}" if base_aup_val else "N/A",
              "Base Units": round(grp_base_units),
              "Incr Units": round(grp_incr_units),
              "Avg Incr Units/Wk": round(grp_avg_incr),
              "Unit Lift %": f"{grp_lift_pct:.1f}%",
          })

        overall_lift_pct = (
            (total_all_incr_units / total_all_base_units * 100)
            if total_all_base_units > 0
            else 0
        )
        avg_incr_per_promo_wk = (
            (total_all_incr_units / promo_weeks_count)
            if promo_weeks_count > 0
            else 0
        )

        st.subheader("=== PROMOTIONAL ANALYSIS REPORT ===")
        st.write(f"• Sheet Analyzed: {sheet_name}")
        st.write(f"• Promo Threshold Used: {int(promo_threshold * 100)}%")
        st.write(f"• Total Weeks Analyzed: {total_weeks}")
        st.write(
            f"• Weeks on Promotion: {promo_weeks_count}"
        )
        st.write(f"• Distinct Promo Groups Identified: {len(groups)}")

        st.subheader("=== PROMO GROUPS BREAKDOWN ===")
        export_df = pd.DataFrame(parsed_groups_records)
        st.dataframe(export_df, use_container_width=True)

        st.subheader("=== TOTAL PROMOTIONAL PERFORMANCE ROLLUP ===")
        st.write(f"• Total Promo Base Units: {total_all_base_units:,.0f}")
        st.write(
            f"• Total Promo Incremental Units: {total_all_incr_units:,.0f}"
        )
        st.write(
            f"• Overall Promotional Unit Lift %: {overall_lift_pct:.1f}%"
        )
        st.write(
            f"• Avg Incr Units / Promo Wk: {avg_incr_per_promo_wk:,.0f}"
        )

        # Download button for CSV export
        csv_data = export_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="Download Analysis as CSV",
            data=csv_data,
            file_name="promotional_analysis_report.csv",
            mime="text/csv",
        )
    except Exception as e:
      st.error(f"Could not parse data file: {e}")