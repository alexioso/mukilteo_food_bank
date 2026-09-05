#TODO: edit google sheets to have filter descending order
# Weekly HH Counts not updating
#First of the month scheduled refresh
#lock down the cells to editor

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
from bb_test import *
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

BROWSERBASE_API_KEY = "bb_live_6Q6UlQnRWatFkdOxKUz3ABabnQM"

bb = Browserbase(api_key=BROWSERBASE_API_KEY)


def save_downloads_on_disk(session_id: str, retry_seconds: int = 20, xlsx=True):
    """
    List and download individual files from a session.
    Retries for the specified number of seconds if no downloads are found.

    :param session_id: The session ID from your browser automation
    :param retry_seconds: How long to retry if no downloads are found
    """
    end_time = time.time() + retry_seconds

    while time.time() < end_time:
        try:
            # List individual downloads for the session
            list_response = requests.get(
                "https://api.browserbase.com/v1/downloads",
                params={"sessionId": session_id},
                headers={"x-bb-api-key": BROWSERBASE_API_KEY},
            )
            data = list_response.json()

            if data["total"] > 0:
                print(f"Found {data['total']} download(s)")

                for download in data["downloads"]:
                    # Download each file individually
                    file_response = requests.get(
                        f"https://api.browserbase.com/v1/downloads/{download['id']}",
                        headers={
                            "x-bb-api-key": BROWSERBASE_API_KEY,
                            "Accept": "application/octet-stream",
                        },
                    )
                    if xlsx:
                        with open("temp.xlsx", "wb") as f:
                            f.write(file_response.content)
                    else:
                        with open("temp.csv", "wb") as f:
                            f.write(file_response.content)
                    print(f"Saved: {download['filename']} ({download['size']} bytes)")
                return
        except Exception as e:
            print(f"Error fetching downloads: {e}")
            raise

        time.sleep(2)  # Wait 2 seconds before retrying

    raise TimeoutError("No downloads found within the retry period")



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
    if "year_month" in result.columns:
        result = result.sort_values("year_month",ascending=False)
    result.to_csv(csv_path, index=False)

    return result

def read_time_entry():


    try:
        print("session")
        session = bb.sessions.create()
        print("connection")
        connection = BrowserbaseConnection(session.id, session.selenium_remote_url)
        options = webdriver.ChromeOptions()
        options.set_capability("se:downloadsEnabled", True)
        
        print("driver")
        driver = webdriver.Remote(
          command_executor=connection, options=options
        )
        print("get url")
        driver.get("https://mfbfp.soxbox.co/login/")
        
        print("username")

        username_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input.rs-input[type="text"]'))
        )

        password_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input.rs-input[type="password"]'))
        )
        
        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Login')]"))
        )


        username_input.send_keys(os.environ["FOOD_BANK_MANAGER_USERNAME"])
        password_input.send_keys(os.environ["FOOD_BANK_MANAGER_PASSWORD"])
        login_button.click()

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
        try:
            save_downloads_on_disk(session.id,xlsx=False)
            print("Downloads complete")
        except Exception as e:
            print(f"Failed to retrieve downloads: {e}")


        df = pd.read_csv("temp.csv")
        
        driver.quit()
        
        return df
    except Exception as e:
        driver.quit()
        raise e    

def read_loaded_report(report_title = 'SCFC Total Report'):
    
    print(f"Exporting: {report_title}")
    

    try:
        #driver = webdriver.Chrome(options=options)
        print("session")
        session = bb.sessions.create()
        print("connection")
        connection = BrowserbaseConnection(session.id, session.selenium_remote_url)

        print("driver")
        options = webdriver.ChromeOptions()
        options.enable_downloads = True
        options.set_capability("se:downloadsEnabled", True) 
        driver = webdriver.Remote(
          command_executor=connection, options=options
        )
        driver.get("https://mfbfp.soxbox.co/login/")
        
        wait = WebDriverWait(driver, 10)
        
        username_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input.rs-input[type="text"]'))
        )

        password_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input.rs-input[type="password"]'))
        )
        
        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Login')]"))
        )


        username_input.send_keys(os.environ["FOOD_BANK_MANAGER_USERNAME"])
        password_input.send_keys(os.environ["FOOD_BANK_MANAGER_PASSWORD"])
        login_button.click()
        

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
        
        try:
            save_downloads_on_disk(session.id)
            print("Downloads complete")
        except Exception as e:
            print(f"Failed to retrieve downloads: {e}")


        df = pd.read_excel("temp.xlsx",skiprows=13)
        
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
    df_total_daily = read_loaded_report("AB: Total Report Daily")

    
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
    df_time_entry_temp = read_time_entry()
    
    #upsert to data/raw paths
    df_total_full = upsert_dataframe(total_report_monthly_path,df_total,"Monthly Visit Date")
    df_dup_full = upsert_dataframe(dup_report_monthly_path,df_dup,"Monthly Visit Date")
    df_undup_full = upsert_dataframe(undup_report_monthly_path,df_undup,"Monthly Visit Date")
    df_time_entry_full = upsert_dataframe(time_entry_path, df_time_entry_temp, "Time Entry ID")
    data_pipeline()
    return 
    

    
import sys
if len(sys.argv) > 1:
    data_pipeline()
else: 
    main_refresh()
