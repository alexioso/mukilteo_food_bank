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

#takes about 3 min to run
def main_refresh():
    df_total = read_loaded_report()
    df_undup = read_loaded_report("SCFC Unduplicated Report")
    df_dup = read_loaded_report("SCFC Duplicated Report")
    df_total_daily = read_loaded_report("AB: Total Report Daily")
    
    #upsert to data/raw paths
    df_total_full = upsert_dataframe(total_report_monthly_path,df_total,"Monthly Visit Date")
    df_dup_full = upsert_dataframe(dup_report_monthly_path,df_dup,"Monthly Visit Date")
    df_undup_full = upsert_dataframe(undup_report_monthly_path,df_undup,"Monthly Visit Date")
    
    
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
    
    upsert_dataframe(total_report_daily_path, df_total_daily, 'date')

    return upsert_dataframe(total_report_daily_path, df_total_daily, 'date')

    
        
main_refresh()