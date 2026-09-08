conda init zsh
conda activate nlp
cd src
python mandatory_report_refresh.py
python generate_forecast.py
#git add ../data/prep/match_stats.csv
git config --global user.name "github-actions[bot]"
git config --global user.email "github-actions[bot]@://github.com"

git add ../data/*
git commit -m "refresh data $(date +%Y-%m-%d)"
git push
#streamlit run dashboard_st.py
