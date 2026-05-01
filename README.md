# World Population Dashboard

Small exploratory data analysis project based on the Kaggle World Population dataset:
https://www.kaggle.com/datasets/sazidthe1/world-population-data

The project cleans country-level population data, explores population trends from 1970 to 2023, and builds an Excel dashboard with continent summaries and country rankings.

## Project Structure

```text
data/
  world_population_data.csv       # original dataset
  world_population_cleaned.csv    # cleaned dataset used for analysis
notebooks/
  project_3.ipynb                 # exploratory analysis
outputs/
  project3_world_population/
    world_population_dashboard.xlsx
scripts/
  build_world_population_dashboard.py
```

## What Is Included

- Cleaning column names and percentage fields
- Population trend analysis by year
- Continent-level summary table
- Top country rankings by population, growth rate, density, and change since 1970
- Excel workbook with a dashboard, charts, source data, and short notes

## How To Run

Create a virtual environment and install dependencies:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Build the dashboard:

```bash
python scripts/build_world_population_dashboard.py
```

The workbook will be saved to:

```text
outputs/project3_world_population/world_population_dashboard.xlsx
```

## Notes

The raw data comes from Kaggle and is included only for educational use. The cleaned file keeps the original country-level records and adds two derived fields:

- `change_since_1970`
- `growth_since_1970`
