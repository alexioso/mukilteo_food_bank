conda init zsh
conda activate nlp
cd src
python mandatory_report_refresh.py 
python google_sheet.py
#git add ../data/prep/match_stats.csv
git add ../data/*
git commit -m "refresh data"
git push
#streamlit run dashboard_st.py
