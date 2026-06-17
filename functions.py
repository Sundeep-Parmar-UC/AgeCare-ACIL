
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
        return (calculate_budget_sum_of_Repair_Maintence(SiteName, LedgerColumnIndex, "Repair", BudgetSheetData) + calculate_budget_sum_of_Repair_Maintence(SiteName, LedgerColumnIndex, "Maintenance", BudgetSheetData) + calculate_budget_sum_of_Other_RnM(SiteName, LedgerColumnIndex, BudgetSheetData))
    elif TargetCategory == "Other (RnM)":
        return calculate_budget_sum_of_Other_RnM(SiteName, LedgerColumnIndex, BudgetSheetData)
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

def calculate_budget_sum_of_Other_RnM(SiteName, LedgerColumnIndex, BudgetSheetData):

    # 1. MATCH logic: Find the target data column using the first 3 characters
    match_str = str(SiteName)[:3].strip().upper()
    header_row = BudgetSheetData.iloc[14] # Excel row 15 is index 14

    target_col_idx = -1
    for col_idx, cell_value in enumerate(header_row):
        if str(cell_value)[:3].strip().upper() == match_str:
            target_col_idx = col_idx
            break

    if target_col_idx == -1:
        return 0.0 # Column header not found

    # 2. SCAN + FILTER logic: Loop through rows to find matching section and sub-category
    total_sum = 0.0
    in_target_section = False

    for i in range(BudgetSheetData.shape[0]):
        try:
            # Clean the Column A value for evaluation
            cell_a_value = str(BudgetSheetData.iloc[i, LedgerColumnIndex]).strip()
            #print("1cell_a_value:",cell_a_value)
            # Replicate the SCAN toggle behavior
            if cell_a_value == "6700:Repair & Maintenance":
                #print("2cell_a_value:",cell_a_value)
                in_target_section = True
                continue
            elif in_target_section and ":" in cell_a_value:
                #print("3cell_a_value:",cell_a_value)
                # Any row with a colon (a new header) turns the toggle off
                in_target_section = False
            elif in_target_section and "Marketing & Advertising" in cell_a_value:
                #print("4cell_a_value:",cell_a_value)
                # Found a title of next section; turn off the switch
                in_target_section = False

            if in_target_section:
              #print("6cell_a_value:",cell_a_value)
              is_other_rnm_category = not (cell_a_value.startswith("Repair") or cell_a_value.startswith("Maintenance"))
              #print("5is_other_rnm_category:",is_other_rnm_category)
              if is_other_rnm_category:
                # Extract value from MonthColumn and add to total, handling non-numeric data
                total_sum += pd.to_numeric(BudgetSheetData.iloc[i, target_col_idx], errors='coerce')
                #print("6total_sum:",total_sum)
                #print("7DataWorkSheet.iloc[i, target_col_idx], i:",i,"  target_col_idx: ",target_col_idx,"  cell_a_value:",cell_a_value)

        except (IndexError, ValueError, TypeError):
            # Gracefully ignore empty rows, text conversion errors, or short rows
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

def calculate_budget_sum_of_Other_RnM_ACIL(LedgerColumnIndex, TargetColumnIndex, BudgetSheetData):
     # 1. SCAN + FILTER logic: Loop through rows to find matching section and sub-category
    total_sum = 0.0
    in_target_section = False

    for i in range(BudgetSheetData.shape[0]):
        try:
            # Clean the Column A value for evaluation
            cell_a_value = str(BudgetSheetData.iloc[i, LedgerColumnIndex]).strip()
            #print("1cell_a_value:",cell_a_value)
            # Replicate the SCAN toggle behavior
            if cell_a_value == "6700:Repair & Maintenance":
                #print("2cell_a_value:",cell_a_value)
                in_target_section = True
                continue
            elif in_target_section and ":" in cell_a_value:
                #print("3cell_a_value:",cell_a_value)
                # Any row with a colon (a new header) turns the toggle off
                in_target_section = False
            elif in_target_section and "Marketing & Advertising" in cell_a_value:
                #print("4cell_a_value:",cell_a_value)
                # Found a title of next section; turn off the switch
                in_target_section = False

            if in_target_section:
              #print("6cell_a_value:",cell_a_value)
              is_other_rnm_category = not (cell_a_value.startswith("Repair") or cell_a_value.startswith("Maintenance"))
              #print("5is_other_rnm_category:",is_other_rnm_category)
              if is_other_rnm_category:
                # Extract value from MonthColumn and add to total, handling non-numeric data
                total_sum += pd.to_numeric(BudgetSheetData.iloc[i, TargetColumnIndex], errors='coerce')
                #print("6total_sum:",total_sum)
                #print("7DataWorkSheet.iloc[i, TargetColumnIndex], i:",i,"  TargetColumnIndex: ",TargetColumnIndex,"  cell_a_value:",cell_a_value)

        except (IndexError, ValueError, TypeError):
            # Gracefully ignore empty rows, text conversion errors, or short rows
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
        return (calculate_budget_sum_of_Repair_Maintence_ACIL(LedgerColumnIndex, TargetColumnIndex,  "Repair", BudgetSheetData) + calculate_budget_sum_of_Repair_Maintence_ACIL(LedgerColumnIndex, TargetColumnIndex,  "Maintenance", BudgetSheetData)+ calculate_budget_sum_of_Other_RnM_ACIL(LedgerColumnIndex, TargetColumnIndex, BudgetSheetData))
    elif TargetCategory == "Other (RnM)":
        return calculate_budget_sum_of_Other_RnM_ACIL(LedgerColumnIndex, TargetColumnIndex, BudgetSheetData)
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


def FindFinCatColumn(DataWorkSheet):
  FinCat_col = 1000
  for r_idx, row in DataWorkSheet.iterrows():
      for c_idx, value in row.items():
          if isinstance(value, str) and value == "Actual":
              if(c_idx < FinCat_col):
                FinCat_col = c_idx
              continue
  return FinCat_col
  
  
def FindSiteColumnFromMonth(DataWorkSheet,Site):
  site_row = -1
  site_col = -1
  for r_idx, row in DataWorkSheet.iterrows():
      for c_idx, value in row.items():
          if isinstance(value, str) and value.strip().upper() == Site.strip().upper():
              site_row = r_idx
              site_col = c_idx
              break
      if site_row != -1:
          break
  return site_row, site_col 

  
def OntarioFundingCheck(FundName):

    FundMapping = {}
    FundMapping['Ontario 4 hours of Care']='Funded'
    FundMapping['Ontario Allied health Professional']='Funded'
    FundMapping['Ontario Behaviour Supports Ontario   (BSO)']='Funded'
    FundMapping['Ontario Behaviour Supports Ontario - OHP/AHP  (BSO)']='Funded'
    FundMapping['Ontario Behaviour Supports Training & Education (BSO Training)']='Funded'
    FundMapping['Ontario Clinical Decision Support funding']='Funded'
    FundMapping['Ontario Convalescent funding']='Funded'
    FundMapping['Ontario Covid Additional PPE funding']='Funded'
    FundMapping['Ontario Covid prevention and containment funding']='Funded'
    FundMapping['Ontario Equipment & Training']='Funded'
    FundMapping['Ontario Fall Prevention Equipment funding']='Funded'
    FundMapping['Ontario Ipac Minor capital']='Funded'
    FundMapping['Ontario Ipac Personnel']='Funded'
    FundMapping['Ontario IPAC Training & Education funding']='Funded'
    FundMapping['Ontario Lab Cost']='Funded'
    FundMapping['Ontario Local Priority Funding']='Funded'
    FundMapping['Ontario LTC Minor capital']='Funded'
    FundMapping['Ontario Medical Safety Technology (MST)']='Funded'
    FundMapping['Ontario Nurse Practioner funding (NP)']='Funded'
    FundMapping['Ontario Nursing and Personal Care (NPC)']='Funded'
    FundMapping['Ontario Nutrional support (NS or previously known as Raw food)']='Funded'
    FundMapping['Ontario Permanent $3 PSW']='Funded'
    FundMapping['Ontario Physician on Call funding']='Funded'
    FundMapping['Ontario Professional growth']='Funded'
    FundMapping['Ontario Program support services (PSS)']='Funded'
    FundMapping['Ontario Resident Health & Wellness']='Funded'
    FundMapping['Ontario Temporary $3 PSW']='Funded'
    FundMapping['Ontario TRIN (Temporary retention incentive for Nurses)']='Funded'
    FundMapping['Palliative Care']='Funded'
    FundMapping['(Blank)']='UnFunded'
    FundMapping['Alberta Baseline Funding']='UnFunded'
    FundMapping['Fundraising and Donation']='UnFunded'
    FundMapping['Ontario HIN']='UnFunded'
    FundMapping['Ontario Other Accomodation (OA)']='UnFunded'
    FundMapping['Private Revenue']='UnFunded'
    FundMapping['Alberta Accommodation Funding']='UnFunded'
    FundMapping['British Columbia Equipcare Funding']='UnFunded'
    FundMapping['Ontario Construction Funding']='Funded'
    FundMapping['$4 premium for NPC Direct care staff - program no longer exist']='UnFunded'

    if FundName in FundMapping:
      return FundMapping[FundName]
    else:
      return "UnFunded"


def OntarioSuppliesCheck(PayrollValue,LedgerValue,RevenueCategory,SpendCategory,BusStream):
    if (PayrollValue == "(Blank)"):
      if(RevenueCategory == "(Blank)"):
        if(BusStream == "Retirement Living"):  # RL
          if(LedgerValue == "6500:Care Supplies"):
            if(SpendCategory == "Nursing supplies"):
              return True
            elif(SpendCategory == "Incontinence"):
              return True

          elif(LedgerValue == "6600:Clinical supplies"):
            if(SpendCategory == "Recreation Supplies"):
              return True

          elif(LedgerValue == "6800:Hospitality"):
            if(SpendCategory == "Kitchen Supplies"):
              return True
            elif(SpendCategory == "Housekeeping - Supplies"):
              return True
            elif(SpendCategory == "Laundry Supplies"):
              return True

        else:  #Bus Stream is LTC
          if(LedgerValue == "6500:Care Supplies"):
            if(SpendCategory == "Nursing supplies"):
              return True
            elif(SpendCategory == "HIN Expense"):
              return True

          elif(LedgerValue == "6600:Clinical supplies"):
            if(SpendCategory == "Recreation Supplies"):
              return True
            elif(SpendCategory == "Eden Initiative"):
              return True

          elif(LedgerValue == "6700:Repair & Maintenance"):
            if(SpendCategory == "Housekeeping - Supplies"):
              return True
            elif(SpendCategory == "Laundry Supplies"):
              return True

          elif(LedgerValue == "6800:Hospitality"):
            if(SpendCategory == "Kitchen Supplies"):
              return True
            elif(SpendCategory == "Housekeeping - Supplies"):
              return True
            elif(SpendCategory == "Laundry Supplies"):
              return True
            elif(SpendCategory == "Kitchen Smallwares"):
              return True
            elif(SpendCategory == "Housekeeping - Linens"):
              return True
            elif(SpendCategory == "Housekeeping - Cleaning Agents"):
              return True
            elif(SpendCategory == "Kitchen Cleaning"):
              return True

    return False


def calculate_sum_of_SUPPLIES(DataWorkSheet, PayrollOrgColumnIndex, LedgerColumnIndex, FundColumnsIndex, BusinessSteamIndex, SpendCategoryIndex, RevenueCategoryIndex, SiteListIndex, ValuesIndex, SiteName):
    filtered_Category_df = DataWorkSheet[
        (DataWorkSheet[SiteListIndex].astype(str).str[0:3] == SiteName)
    ]
    total_sum = 0
    #walk through all values in the FundColumnindex,  if it maps to Unfunded then check if Supplies is true then add the Valueindex column to total sum
    for index, row in filtered_Category_df.iterrows():
      if OntarioFundingCheck(row[FundColumnsIndex]) == "UnFunded":
          if OntarioSuppliesCheck(row[PayrollOrgColumnIndex],row[LedgerColumnIndex],row[RevenueCategoryIndex],row[SpendCategoryIndex],row[BusinessSteamIndex]):
              total_sum += row[ValuesIndex]

    return total_sum


def calculate_sum_by_category_IrisRecency(DataWorkSheet, SpendCategoryIndex, SiteListIndex, ValuesIndex, SiteName, Category):
    filtered_Category_df = DataWorkSheet[
        (DataWorkSheet[SpendCategoryIndex].astype(str) == Category)
    ]
    filtered_Site_Cat_df = filtered_Category_df[
        (filtered_Category_df[SiteListIndex].astype(str).str[0:3] == SiteName)
    ]
    total_sum = filtered_Site_Cat_df[ValuesIndex].apply(pd.to_numeric, errors='coerce').sum()
    return total_sum


def calculate_sum_by_Fund_Unfund_IrisRecency(DataWorkSheet, LedgerColumnIndex, FundColumnsIndex, SiteListIndex, ValuesIndex, SiteName, Category):
    filtered_Category_df = DataWorkSheet[
        (DataWorkSheet[LedgerColumnIndex].astype(str) == "6700:Repair & Maintenance")
    ]
    filtered_Site_Cat_df = filtered_Category_df[
        (filtered_Category_df[SiteListIndex].astype(str).str[0:3] == SiteName)
    ]
    total_sum = 0
    #walk through all values in the FundColumnindex,  if it maps to Category then add the Valueindex column to total sum
    for index, row in filtered_Site_Cat_df.iterrows():
      if OntarioFundingCheck(row[FundColumnsIndex]) == Category:
        total_sum += row[ValuesIndex]

    return total_sum    


def calculate_value_IrisRecency(DataWorkSheet, PayrollOrgColumnIndex, LedgerColumnIndex, FundColumnsIndex, BusinessSteamIndex, SpendCategoryIndex, RevenueCategoryIndex, SiteListIndex, ValuesIndex, SiteName, Category):
    StraightSumCategory = ["Sick Time", "Overtime"]
    if Category in StraightSumCategory:
        return calculate_sum_by_category_IrisRecency(DataWorkSheet, SpendCategoryIndex, SiteListIndex, ValuesIndex, SiteName, Category)
    elif Category == "Purchased Hours" or Category == "Purchased Hours total":
        NewSumTotal = calculate_sum_by_category_IrisRecency(DataWorkSheet, SpendCategoryIndex, SiteListIndex, ValuesIndex, SiteName, "Purchased Hours")
        NewPodTotal = calculate_sum_by_category_IrisRecency(DataWorkSheet, RevenueCategoryIndex, SiteListIndex, ValuesIndex, SiteName, "Podiatry")
        return NewSumTotal + NewPodTotal
    elif Category == "6500:Care Supplies":
      return calculate_sum_by_category_IrisRecency(DataWorkSheet, LedgerColumnIndex, SiteListIndex, ValuesIndex, SiteName, "6500:Care Supplies") + calculate_sum_by_category_IrisRecency(DataWorkSheet, LedgerColumnIndex, SiteListIndex, ValuesIndex, SiteName, "6600:Clinical supplies")
    elif Category == "Repair":
        return calculate_sum_by_Fund_Unfund_IrisRecency(DataWorkSheet, LedgerColumnIndex, FundColumnsIndex, SiteListIndex, ValuesIndex, SiteName, "Funded")
    elif Category == "Maintenance":
        return calculate_sum_by_Fund_Unfund_IrisRecency(DataWorkSheet, LedgerColumnIndex, FundColumnsIndex, SiteListIndex, ValuesIndex, SiteName, "UnFunded")
    elif Category == "SUPPLIES":
        return calculate_sum_of_SUPPLIES(DataWorkSheet, PayrollOrgColumnIndex, LedgerColumnIndex, FundColumnsIndex, BusinessSteamIndex, SpendCategoryIndex, RevenueCategoryIndex, SiteListIndex, ValuesIndex, SiteName)
    else:
        return 0


def calculate_sum_Overtime_Budget_IrisRecency(DataWorkSheet, LedgerColumnIndex, SiteListIndex, ValuesIndex, SiteName):
    filtered_Category_df = DataWorkSheet[
        (DataWorkSheet[LedgerColumnIndex].astype(str) == "6000:Salaries and wages")
    ]
    filtered_Site_Cat_df = filtered_Category_df[
        (filtered_Category_df[SiteListIndex].astype(str).str[0:3] == SiteName)
    ]
    total_sum = filtered_Site_Cat_df[ValuesIndex].apply(pd.to_numeric, errors='coerce').sum()
    return total_sum


def calculate_value_Budget_IrisRecency(DataWorkSheet, PayrollOrgColumnIndex, LedgerColumnIndex, FundColumnsIndex, BusinessSteamIndex, SpendCategoryIndex, RevenueCategoryIndex, SiteListIndex, ValuesIndex, SiteName, Category):
    if Category == "Sick Time":
        return calculate_sum_by_category_IrisRecency(DataWorkSheet, SpendCategoryIndex, SiteListIndex, ValuesIndex, SiteName, Category)
    elif Category == "Overtime":
        return 0.012*calculate_sum_Overtime_Budget_IrisRecency(DataWorkSheet, LedgerColumnIndex, SiteListIndex, ValuesIndex, SiteName)
    elif Category == "Purchased Hours" or Category == "Purchased Hours total":
        return calculate_sum_by_category_IrisRecency(DataWorkSheet, SpendCategoryIndex, SiteListIndex, ValuesIndex, SiteName, "Purchased Hours")
    elif Category == "6500:Care Supplies":
      return calculate_sum_by_category_IrisRecency(DataWorkSheet, LedgerColumnIndex, SiteListIndex, ValuesIndex, SiteName, "6500:Care Supplies") + calculate_sum_by_category_IrisRecency(DataWorkSheet, LedgerColumnIndex, SiteListIndex, ValuesIndex, SiteName, "6600:Clinical supplies")
    elif Category == "Repair":
        return calculate_sum_by_Fund_Unfund_IrisRecency(DataWorkSheet, LedgerColumnIndex, FundColumnsIndex, SiteListIndex, ValuesIndex, SiteName, "Funded")
    elif Category == "Maintenance":
        return calculate_sum_by_Fund_Unfund_IrisRecency(DataWorkSheet, LedgerColumnIndex, FundColumnsIndex, SiteListIndex, ValuesIndex, SiteName, "UnFunded")
    elif Category == "SUPPLIES":
        return calculate_sum_of_SUPPLIES(DataWorkSheet, PayrollOrgColumnIndex, LedgerColumnIndex, FundColumnsIndex, BusinessSteamIndex, SpendCategoryIndex, RevenueCategoryIndex, SiteListIndex, ValuesIndex, SiteName)
    else:
        return 0
    
    