import os
from flask import Flask, render_template, redirect, request
import pandas as pd

app = Flask(__name__)

# Milestone thresholds
MINIFIG_THRESHOLDS = [
    (150, "Grand Master"),
    (80, "Platinum"),
    (40, "Gold"),
    (20, "Silver"),
    (10, "Bronze"),
    (0, "Beginner")
]
POINTS_THRESHOLDS = [
    (800, "Grand Master"),
    (400, "Platinum"),
    (200, "Gold"),
    (100, "Silver"),
    (50, "Bronze"),
    (0, "Beginner")
]

def get_tier(value, thresholds):
    for thresh, name in thresholds:
        if value >= thresh:
            return name
    return thresholds[-1][1]

# Load Excel
def load_data():
    df = pd.read_excel("hall_of_fame.xlsx")

    if 'Total_Points' in df.columns:
        df['Total_Points'] = pd.to_numeric(df['Total_Points'], errors='coerce').fillna(0).astype(int)
    else:
        df['Total_Points'] = 0

    if 'Total_Minifigs' in df.columns:
        df['Total_Minifigs'] = pd.to_numeric(df['Total_Minifigs'], errors='coerce').fillna(0).astype(int)
    else:
        df['Total_Minifigs'] = 0

    df = df.fillna("—")

    df['Points_Tier'] = df['Total_Points'].apply(lambda v: get_tier(v, POINTS_THRESHOLDS))
    df['Minifig_Tier'] = df['Total_Minifigs'].apply(lambda v: get_tier(v, MINIFIG_THRESHOLDS))

    tier_order = ["Beginner", "Bronze", "Silver", "Gold", "Platinum", "Grand Master"]
    rank = {t: i for i, t in enumerate(tier_order)}

    def overall_tier(row):
        p = row['Points_Tier']
        m = row['Minifig_Tier']
        return p if rank[p] >= rank[m] else m

    df['Overall_Tier'] = df.apply(overall_tier, axis=1)

    def tier_source(row):
        return 'Points' if rank[row['Points_Tier']] >= rank[row['Minifig_Tier']] else 'Minifigs'

    df['Tier_Source'] = df.apply(tier_source, axis=1)

    return df

@app.route('/')
def home():
    df = load_data().sort_values(by='Total_Points', ascending=False)
    top3 = df.head(3).to_dict(orient='records')
    return render_template('home.html', top3=top3)

@app.route('/top-collectors')
def top_collectors():
    df = load_data().sort_values(by='Total_Points', ascending=False)
    data = df.to_dict(orient='records')
    return render_template('top_collectors.html', collectors=data)

@app.route('/submit')
def submit():
    return redirect("https://forms.gle/5wYMhzYjZS9vY9Sv8")

@app.route('/about')
def about():
    return render_template('about.html')

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
