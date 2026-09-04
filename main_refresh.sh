conda init zsh
conda activate nlp
cd src
python mandatory_report_refresh.py
python generate_forecast.py
#git add ../data/prep/match_stats.csv
git add ../data/*
git commit -m "refresh data $(date +%Y-%m-%d)"
git push
#streamlit run dashboard_st.py
