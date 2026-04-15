import pandas as pd
from dotenv import load_dotenv
import os
load_dotenv()
import requests
from datetime import date
from calendar import monthrange
import time
import glob
from config import *

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

def read_most_recent_csv(folder_path: str) -> pd.DataFrame:
    csv_files = glob.glob(os.path.join(folder_path, "*.csv"))
    if not csv_files:
        raise FileNotFoundError(f"No CSV files found in {folder_path}")

    most_recent = max(csv_files, key=os.path.getmtime)
    print(f"Reading: {most_recent}")
    return pd.read_csv(most_recent)

def read_most_recent_xlsx(folder_path: str,
                          skip_rows: int = 13) -> pd.DataFrame:
    csv_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    if not csv_files:
        raise FileNotFoundError(f"No xlsx files found in {folder_path}")

    most_recent = max(csv_files, key=os.path.getmtime)
    print(f"Reading: {most_recent}")
    return pd.read_excel(most_recent,skiprows=skip_rows)

def get_chrome_default_download_path():
    if os.name == 'nt':  # Windows
        return os.path.join(os.environ['USERPROFILE'], 'Downloads')
    elif os.name == 'posix':
        if os.uname().sysname == 'Darwin':  # macOS
            return os.path.join(os.path.expanduser('~'), 'Downloads')
        else:  # Linux
            return os.path.join(os.path.expanduser('~'), 'Downloads')

def get_prior_month_range(current_date: date) -> tuple[int, int, int, int]:
    """
    Returns (year, month, day0, day1) for the first and last day of the prior month.
    """
    # Go to first day of current month, then back one day
    first_of_current = current_date.replace(day=1)
    last_of_prior = first_of_current.replace(day=1) - __import__('datetime').timedelta(days=1)

    year = last_of_prior.year
    month = last_of_prior.month
    day0 = 1
    day1 = monthrange(year, month)[1]  # last day of that month

    return year, month, day0, day1

def get_distribution_report_url(anchor_date):
    
    year, month, day0, day1 = get_prior_month_range(anchor_date)
    
    if len(str(month)) == 1:
        month = "0" + str(month)
    else:
        month = str(month)
    day0 = "01"
    url = f"https://mfbfp.soxbox.co/reports/outreach/outreach-details/?startPicker={month}%2F{day0}%2F{year}&endPicker={month}%2F{day1}%2F{year}#export"
    return(url)


def upsert_dataframe(csv_path: str, new_rows: pd.DataFrame, key_columns: list) -> pd.DataFrame:
    """
    Upserts new_rows into the DataFrame loaded from csv_path.
    
    Args:
        csv_path:    Path to the CSV file containing existing data
        new_rows:    DataFrame with new/updated rows (same schema)
        key_columns: List of column names to use as the composite key
    
    Returns:
        Updated DataFrame with upserted rows (also saves back to CSV)
    """
    existing = pd.read_csv(csv_path)

    # Set key columns as index for both DataFrames
    existing_indexed = existing.set_index(key_columns)
    new_indexed = new_rows.set_index(key_columns)

    # Update existing rows and append new ones
    existing_indexed = existing_indexed.combine_first(new_indexed)  # adds new keys
    existing_indexed.update(new_indexed)                            # updates existing keys

    result = existing_indexed.reset_index()

    # Save back to CSV
    result.to_csv(csv_path, index=False)

    return result

def read_time_entry():
    options = Options()
    options.add_argument("--headless")

    try:
        driver = webdriver.Chrome(options=options)
        driver.get("https://mfbfp.soxbox.co/login/")

        driver.find_element(By.NAME, "username").send_keys(os.environ["FOOD_BANK_MANAGER_USERNAME"])
        driver.find_element(By.NAME, "password").send_keys(os.environ["FOOD_BANK_MANAGER_PASSWORD"])
        driver.find_element(By.NAME, "action").click()

        time.sleep(2)

        #old report
        #report_url = get_distribution_report_url(anchor_date)
        report_url = "https://mfbfp.soxbox.co/reports/team/time-entry/"
        driver.get(report_url)

        time.sleep(2)

        #click load button to the SCFC Total Report
        search_button = driver.find_element(By.CSS_SELECTOR, "input[type='submit'][value='Search']")
        driver.execute_script("arguments[0].click();", search_button)


        #wait for export button and then click
        export_button = WebDriverWait(driver, 120).until(
            EC.element_to_be_clickable((
                By.LINK_TEXT, "Export to CSV"
            ))
        )
        export_button.click()
        
        time.sleep(5)

        df = read_most_recent_csv(get_chrome_default_download_path())
        
        driver.quit()
        
        return df
    except Exception as e:
        driver.quit()
        raise e    

def read_loaded_report(report_title = 'SCFC Total Report'):
    
    print(f"Exporting: {report_title}")
    
    options = Options()
    options.add_argument("--headless")

    try:
        driver = webdriver.Chrome(options=options)
        driver.get("https://mfbfp.soxbox.co/login/")

        driver.find_element(By.NAME, "username").send_keys(os.environ["FOOD_BANK_MANAGER_USERNAME"])
        driver.find_element(By.NAME, "password").send_keys(os.environ["FOOD_BANK_MANAGER_PASSWORD"])
        driver.find_element(By.NAME, "action").click()

        time.sleep(2)

        #old report
        #report_url = get_distribution_report_url(anchor_date)
        report_url = "https://mfbfp.soxbox.co/reports/outreach/outreach-aggregate"
        driver.get(report_url)

        time.sleep(2)

        #load the preloaded reports
        driver.find_element(By.XPATH, "//h5[contains(text(), 'Preset: OAR Report')]").click()
        time.sleep(2)
        #click load button to the SCFC Total Report
        load_button = driver.find_element(
            By.XPATH,
            f"//div[@role='row'][.//div[contains(@class,'rs-table-cell-content') and text()='{report_title}']]//button[contains(text(),'Load')]"
        )
        driver.execute_script("arguments[0].click();", load_button)

        #click search button
        button = driver.find_element(By.XPATH, "//button[contains(@class,'rs-btn-primary') and contains(text(),'Search')]")
        button.click()

        #wait for export button and then click
        export_button = WebDriverWait(driver, 120).until(
            EC.element_to_be_clickable((
                By.XPATH, "//button[contains(@class,'rs-btn-ghost') and contains(text(),'Export')]"
            ))
        )
        export_button.click()
        
        time.sleep(5)

        df = read_most_recent_xlsx(get_chrome_default_download_path())
        
        driver.quit()
        
        return df
    except Exception as e:
        driver.quit()
        raise e
    
    
def data_pipeline():
    df_total_full = pd.read_csv(total_report_monthly_path)
    df_dup_full = pd.read_csv(dup_report_monthly_path)
    df_undup_full = pd.read_csv(undup_report_monthly_path)
    df_time_entry_full = pd.read_csv(time_entry_path)
    df_total_daily = pd.read_csv(total_report_daily_path)
    
    #combine dataframes
    df_monthly = df_total_full.rename({'# of HH Visits':'total_hh_visits', 
                    '0 to 2':'total_age_0_to_2', 
                    '3 to 18':'total_age_3_to_18', 
                    '19 to 54':'total_age_19_to_54', 
                    '55+':'total_age_55_plus',
                    'Age: Not Provided' : 'total_age_anonymous', 
                    'Total Individuals' : 'total_indivdiduals', 
                    'Monthly Visit Date' : 'year_month',
                    'Total weight' : 'total_weight'
                    },axis=1).merge(df_dup_full.rename({'# of HH Visits':'dup_hh_visits', 
                    '0 to 2':'dup_age_0_to_2', 
                    '3 to 18':'dup_age_3_to_18', 
                    '19 to 54':'dup_age_19_to_54', 
                    '55+':'dup_age_55_plus',
                    'Age: Not Provided' : 'dup_age_anonymous', 
                    'Total Individuals' : 'dup_indivdiduals', 
                    'Monthly Visit Date' : 'year_month',
                    'Total weight' : 'dup_weight'
                    },axis=1), how = 'outer', on = 'year_month').\
                    merge(df_undup_full.rename({'# of HH Visits':'undup_hh_visits', 
                    '0 to 2':'undup_age_0_to_2', 
                    '3 to 18':'undup_age_3_to_18', 
                    '19 to 54':'undup_age_19_to_54', 
                    '55+':'undup_age_55_plus',
                    'Age: Not Provided' : 'undup_age_anonymous', 
                    'Total Individuals' : 'undup_indivdiduals', 
                    'Monthly Visit Date' : 'year_month',
                    'Total weight' : 'undup_weight'
                    },axis=1), how = 'outer', on = 'year_month')
                    
 
    df_time_entry_full["Time Entry On"] = pd.to_datetime(df_time_entry_full["Time Entry On"])
    df_time_entry_full["year_month"] = df_time_entry_full["Time Entry On"].dt.year.astype(str) + "-" + df_time_entry_full["Time Entry On"].dt.month.astype(str).str.zfill(2)

    df_time_entry_grouped = df_time_entry_full.\
        groupby("year_month").agg(
            volunteer_count=("Volunteer ID", "nunique"),
            volunteer_set=("Volunteer ID", set),
            volunteer_hours=("Hours Worked","sum")
            ).reset_index()
            
    df_time_entry_grouped["month"] = df_time_entry_grouped["year_month"].str.split("-").str[1].astype(int)

    df_time_entry_grouped["volunteer_hours_ytd"] = df_time_entry_grouped.apply(
        lambda row: df_time_entry_grouped.loc[:row.name, "volunteer_hours"].tail(row["month"]).sum(),
        axis=1
    )

    #rolling sum of set col for total volunteers YTD
    results = []
    for i, row in df_time_entry_grouped.iterrows():
        window = row["month"]
        start = max(0, i - window + 1)
        # Filter out NaN values before union
        sets_in_window = [s for s in df_time_entry_grouped.loc[start:i, "volunteer_set"] if isinstance(s, set)]
        union_set = set().union(*sets_in_window) if sets_in_window else set()
        results.append(union_set)

    df_time_entry_grouped["volunteer_set_ytd"] = results

    df_time_entry_grouped["volunteer_count_ytd"] = df_time_entry_grouped["volunteer_set_ytd"].apply(lambda x: len(x) if isinstance(x, set) else 0)
        
    df_monthly = df_monthly.merge(df_time_entry_grouped[["year_month","volunteer_count","volunteer_hours","volunteer_count_ytd","volunteer_hours_ytd"]],how="left",on="year_month")
                    
    df_total_daily = df_total_daily.rename({'# of HH Visits':'total_hh_visits', 
                    '0 to 2':'total_age_0_to_2', 
                    '3 to 18':'total_age_3_to_18', 
                    '19 to 54':'total_age_19_to_54', 
                    '55+':'total_age_55_plus',
                    'Age: Not Provided' : 'total_age_anonymous', 
                    'Total Individuals' : 'total_indivdiduals', 
                    'Visit Date' : 'date',
                    'Total weight' : 'total_weight'
                    },axis=1)
    
    df_total_daily['week_start'] = pd.to_datetime(df_total_daily['date']).dt.to_period('W').apply(lambda r: r.start_time).dt.strftime("%Y-%m-%d")
    df_total_daily["day_of_week"] = pd.to_datetime(df_total_daily["date"]).dt.day_name()
    df_total_daily["month"] = pd.to_datetime(df_total_daily["date"]).dt.month_name()
    df_weekly = df_total_daily.pivot_table(index="week_start", columns="day_of_week", values="total_hh_visits", aggfunc="sum").reset_index()
    df_weekly = df_weekly.loc[:,["week_start"]+weekly_days_of_week]
    df_weekly["Total"] = df_weekly[weekly_days_of_week].sum(axis=1)
    upsert_dataframe(df_weekly_path, df_weekly[["week_start"] + weekly_days_of_week + ["Total"]], 'week_start')
    
    upsert_dataframe(df_monthly_path, df_monthly, 'year_month')

    return upsert_dataframe(total_report_daily_path, df_total_daily, 'date')

#takes about 3 min to run
def main_refresh():
    df_total = read_loaded_report()
    df_undup = read_loaded_report("SCFC Unduplicated Report")
    df_dup = read_loaded_report("SCFC Duplicated Report")
    df_total_daily = read_loaded_report("AB: Total Report Daily")
    df_time_entry_temp = read_time_entry()
    
    #upsert to data/raw paths
    df_total_full = upsert_dataframe(total_report_monthly_path,df_total,"Monthly Visit Date")
    df_dup_full = upsert_dataframe(dup_report_monthly_path,df_dup,"Monthly Visit Date")
    df_undup_full = upsert_dataframe(undup_report_monthly_path,df_undup,"Monthly Visit Date")
    df_time_entry_full = upsert_dataframe(time_entry_path, df_time_entry_temp, "Time Entry ID")
    
    return data_pipeline()
    

    
import sys
if len(sys.argv) > 1:
    data_pipeline()
else: 
    main_refresh()