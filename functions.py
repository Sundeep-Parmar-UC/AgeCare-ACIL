
import pandas as pd
import re
import colorsys
import numpy as np

def FindLedgerColumn(DataWorkSheet):
  ledger_account_col = 1000
  for r_idx, row in DataWorkSheet.iterrows():
      for c_idx, value in row.items():
          if isinstance(value, str) and value == "Ledger Account":
              if(c_idx < ledger_account_col):
                ledger_account_col = c_idx
              continue
  return ledger_account_col

def FindMonthActualsColumn(DataWorkSheet,SummaryMonth):
  month_row = -1
  month_col = -1
  for r_idx, row in DataWorkSheet.iterrows():
      for c_idx, value in row.items():
          if isinstance(value, str) and value.strip().upper() == SummaryMonth.strip().upper():
              month_row = r_idx
              month_col = c_idx
              break
      if month_row != -1:
          break
  return month_row, month_col

def calculate_sum_by_category(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex, Category):
    filtered_df = DataWorkSheet[
        (DataWorkSheet[LedgerColumnIndex].astype(str) == Category)
    ]
    total_sum = filtered_df[MonthColumnIndex].apply(pd.to_numeric, errors='coerce').sum()
    return total_sum

def calculate_value(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex, Category):
    StraightSumCategory = ["Sick Time", "Overtime", "Purchased Hours", "6500:Care Supplies", "6600:Clinical supplies", "Nursing Supplies - Added Care Expense", "Added-Care", "6700:Repair & Maintenance"]
    if Category in StraightSumCategory:
        return calculate_sum_by_category(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex, Category)
    elif Category == "Purchased Hours total":
        return (calculate_value(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex, "Purchased Hours") - calculate_value(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex, "- Added Care"))
    elif Category == "SUPPLIES":
        return (calculate_value(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex, "6500:Care Supplies") + calculate_value(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex, "6600:Clinical supplies") - calculate_value(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex, "Nursing Supplies - Added Care Expense"))
    elif Category == "SUM":
        return (calculate_value(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex, "Repair") + calculate_value(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex, "Maintenance") + calculate_value(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex, "Other (RnM)"))
    elif Category == "- Added Care":
        return calculate_sum_of_Added_Care(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex)
    elif Category == "Repair" or Category == "Maintenance":
        return calculate_sum_of_Repair_Maintence(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex, Category)
    elif Category == "Other (RnM)":
        return calculate_sum_of_Other_RnM(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex, Category)
    else:
        return -1000000

def calculate_sum_of_Added_Care(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex):
    total_added_care_sum = 0
    for i in range(DataWorkSheet.shape[0] - 1):
        try:
            current_row_val = str(DataWorkSheet.iloc[i, LedgerColumnIndex])
            next_row_val = str(DataWorkSheet.iloc[i + 1, LedgerColumnIndex])
            if current_row_val == "Added Care HCA" and next_row_val == "Purchased Hours":
                total_added_care_sum += pd.to_numeric(DataWorkSheet.iloc[i + 1, MonthColumnIndex], errors='coerce')
        except IndexError:
            continue
        except ValueError:
            continue
    return total_added_care_sum

def calculate_sum_of_Repair_Maintence(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex, Category):
    total_sum = 0
    in_target_section = False
    for i in range(DataWorkSheet.shape[0]):
        try:
            cell_a_value = str(DataWorkSheet.iloc[i, LedgerColumnIndex]).strip()
            if cell_a_value == "6700:Repair & Maintenance":
                in_target_section = True
            elif in_target_section and ":" in cell_a_value:
                in_target_section = False
            starts_with_repair = cell_a_value.startswith(Category)
            if in_target_section and starts_with_repair:
                total_sum += pd.to_numeric(DataWorkSheet.iloc[i, MonthColumnIndex], errors='coerce')
        except (IndexError, ValueError, TypeError):
            continue
    return total_sum

def calculate_sum_of_Other_RnM(DataWorkSheet, LedgerColumnIndex, MonthColumnIndex, Category):
    total_sum = 0
    in_target_section = False
    for i in range(DataWorkSheet.shape[0]):
        try:
            cell_a_value = str(DataWorkSheet.iloc[i, LedgerColumnIndex]).strip()
            if cell_a_value == "6700:Repair & Maintenance":
                in_target_section = True
                continue
            elif in_target_section and ":" in cell_a_value:
                in_target_section = False
                continue
            elif in_target_section and "Marketing & Advertising" in cell_a_value:
                in_target_section = False
                continue
            is_other_rnm_category = not (cell_a_value.startswith("Repair") or cell_a_value.startswith("Maintenance"))
            if in_target_section and is_other_rnm_category:
                total_sum += pd.to_numeric(DataWorkSheet.iloc[i, MonthColumnIndex], errors='coerce')
        except (IndexError, ValueError, TypeError):
            continue
    return total_sum

def calculate_budget_sum_general(SiteName, LedgerColumnIndex, TargetCategory, BudgetSheetData):
    header_row = BudgetSheetData.iloc[14]
    target_column_index = -1
    for col_idx, cell_value in enumerate(header_row):
        if str(cell_value)[:3].strip().upper() == SiteName:
            target_column_index = col_idx
            break
    if target_column_index == -1:
        return 0.0
    total_sum = 0.0
    for i in range(BudgetSheetData.shape[0]):
        try:
            if str(BudgetSheetData.iloc[i, LedgerColumnIndex]).strip() == str(TargetCategory).strip():
                total_sum += pd.to_numeric(BudgetSheetData.iloc[i, target_column_index], errors='coerce')
        except (IndexError, ValueError, TypeError):
            continue
    return total_sum

def calculate_budget_value(SiteName, LedgerColumnIndex, TargetCategory, BudgetSheetData):
    StraightBudgetSumCategory = ["Sick Time", "Overtime", "Purchased Hours", "6500:Care Supplies", "6600:Clinical supplies", "6700:Repair & Maintenance"]
    if TargetCategory in StraightBudgetSumCategory:
        return calculate_budget_sum_general(SiteName, LedgerColumnIndex, TargetCategory, BudgetSheetData)
    elif TargetCategory == "Repair" or TargetCategory == "Maintenance":
        return calculate_budget_sum_of_Repair_Maintence(SiteName, LedgerColumnIndex, TargetCategory, BudgetSheetData)
    elif TargetCategory == "SUPPLIES":
        return (calculate_budget_sum_general(SiteName, LedgerColumnIndex, "6500:Care Supplies", BudgetSheetData) + calculate_budget_sum_general(SiteName, LedgerColumnIndex, "6600:Clinical supplies", BudgetSheetData))
    elif TargetCategory == "SUM":
        return (calculate_budget_sum_of_Repair_Maintence(SiteName, LedgerColumnIndex, "Repair", BudgetSheetData) + calculate_budget_sum_of_Repair_Maintence(SiteName, LedgerColumnIndex, "Maintenance", BudgetSheetData))
    else:
        return -1000000

def calculate_budget_sum_of_Repair_Maintence(SiteName, LedgerColumnIndex, TargetCategory, BudgetSheetData):
    match_str = str(SiteName)[:3].strip().upper()
    header_row = BudgetSheetData.iloc[14]
    target_col_idx = -1
    for col_idx, cell_value in enumerate(header_row):
        if str(cell_value)[:3].strip().upper() == match_str:
            target_col_idx = col_idx
            break
    if target_col_idx == -1:
        return 0.0
    total_sum = 0.0
    in_target_section = False
    sub_cat_str = str(TargetCategory).strip()
    sub_cat_len = len(sub_cat_str)
    for i in range(BudgetSheetData.shape[0]):
        try:
            cell_a_value = str(BudgetSheetData.iloc[i, LedgerColumnIndex]).strip()
            if cell_a_value == "6700:Repair & Maintenance":
                in_target_section = True
            elif in_target_section and ":" in cell_a_value:
                in_target_section = False
            starts_with_sub_cat = cell_a_value[:sub_cat_len] == sub_cat_str if sub_cat_len > 0 else False
            if in_target_section and starts_with_sub_cat:
                total_sum += pd.to_numeric(BudgetSheetData.iloc[i, target_col_idx], errors='coerce')
        except (IndexError, ValueError, TypeError):
            continue
    return total_sum

def calculate_budget_sum_general_ACIL(LedgerColumnIndex, TargetColumnIndex, TargetCategory, BudgetSheetData):
    total_sum = 0.0
    for i in range(BudgetSheetData.shape[0]):
        try:
            if str(BudgetSheetData.iloc[i, LedgerColumnIndex]).strip() == str(TargetCategory).strip():
                total_sum += pd.to_numeric(BudgetSheetData.iloc[i, TargetColumnIndex], errors='coerce')
        except (IndexError, ValueError, TypeError):
            continue
    return total_sum

def calculate_budget_sum_of_Repair_Maintence_ACIL(LedgerColumnIndex, TargetColumnIndex, TargetCategory, BudgetSheetData):
    total_sum = 0.0
    in_target_section = False
    sub_cat_str = str(TargetCategory).strip()
    sub_cat_len = len(sub_cat_str)
    for i in range(BudgetSheetData.shape[0]):
        try:
            cell_a_value = str(BudgetSheetData.iloc[i, LedgerColumnIndex]).strip()
            if cell_a_value == "6700:Repair & Maintenance":
                in_target_section = True
            elif in_target_section and ":" in cell_a_value:
                in_target_section = False
            starts_with_sub_cat = cell_a_value[:sub_cat_len] == sub_cat_str if sub_cat_len > 0 else False
            if in_target_section and starts_with_sub_cat:
                total_sum += pd.to_numeric(BudgetSheetData.iloc[i, TargetColumnIndex], errors='coerce')
        except (IndexError, ValueError, TypeError):
            continue
    return total_sum

def calculate_budget_value_ACIL(LedgerColumnIndex, TargetColumnIndex, TargetCategory, BudgetSheetData):
    StraightBudgetSumCategory = ["Sick Time", "Overtime", "Purchased Hours", "6500:Care Supplies", "6600:Clinical supplies", "6700:Repair & Maintenance"]
    if TargetCategory in StraightBudgetSumCategory:
        return calculate_budget_sum_general_ACIL(LedgerColumnIndex, TargetColumnIndex, TargetCategory, BudgetSheetData)
    elif TargetCategory == "Repair" or TargetCategory == "Maintenance":
        return calculate_budget_sum_of_Repair_Maintence_ACIL(LedgerColumnIndex, TargetColumnIndex, TargetCategory, BudgetSheetData)
    elif TargetCategory == "SUPPLIES":
        return (calculate_budget_sum_general_ACIL(LedgerColumnIndex, TargetColumnIndex,  "6500:Care Supplies", BudgetSheetData) + calculate_budget_sum_general_ACIL(LedgerColumnIndex, TargetColumnIndex,  "6600:Clinical supplies", BudgetSheetData))
    elif TargetCategory == "SUM":
        return (calculate_budget_sum_of_Repair_Maintence_ACIL(LedgerColumnIndex, TargetColumnIndex,  "Repair", BudgetSheetData) + calculate_budget_sum_of_Repair_Maintence_ACIL(LedgerColumnIndex, TargetColumnIndex,  "Maintenance", BudgetSheetData))
    else:
        return -1000000

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def rgb_to_hex(rgb_color):
    return '#%02x%02x%02x' % rgb_color

def interpolate_color(start_hex, end_hex, factor):
    start_rgb = hex_to_rgb(start_hex)
    end_rgb = hex_to_rgb(end_hex)
    interpolated_rgb = tuple(
        int(start_rgb[i] + factor * (end_rgb[i] - start_rgb[i])) for i in range(3)
    )
    return rgb_to_hex(interpolated_rgb)

def get_contrasting_font_color(bg_hex_color):
    r, g, b = hex_to_rgb(bg_hex_color)
    r /= 255.0
    g /= 255.0
    b /= 255.0
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    return '#000000' if l > 0.5 else '#FFFFFF'
