"""
Netflix Titles – quick analysis (runs in ~1–2 minutes)

What it does:
- Loads netflix_titles.xlsx
- Cleans duration into numeric minutes/seasons
- Creates 3 simple charts and saves them as PNGs
- Writes a small summary table to CSV

How to run (Windows):
1) Put this file in the same folder as netflix_titles.xlsx
2) Open Command Prompt in that folder
3) Run:  python netflix_analysis.py

Requirements:
- Python 3.9+
- pandas, matplotlib, openpyxl
Install once:
    pip install pandas matplotlib openpyxl
"""

from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt

HERE = Path(__file__).resolve().parent
DATA_FILE = HERE / "netflix_titles.xlsx"
OUT_DIR = HERE / "outputs"
OUT_DIR.mkdir(exist_ok=True)

def pick_title_sheet(excel_path: Path) -> pd.DataFrame:
    """Pick the best sheet that contains the main title-level table."""
    xls = pd.ExcelFile(excel_path)
    preferred = ["netflix_titles", "titles", "sheet1", "Sheet1"]
    for name in preferred:
        if name in xls.sheet_names:
            return pd.read_excel(excel_path, sheet_name=name)
    return pd.read_excel(excel_path, sheet_name=xls.sheet_names[0])

def main():
    if not DATA_FILE.exists():
        raise FileNotFoundError(
            f"Could not find {DATA_FILE.name} in:\n{HERE}\n"
            "Put netflix_titles.xlsx in the same folder as this script."
        )

    df = pick_title_sheet(DATA_FILE)

    # Standardize column names
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]

    # Parse duration if needed
    if "duration" in df.columns and ("duration_minutes" not in df.columns and "duration_seasons" not in df.columns):
        dur = df["duration"].astype(str)
        df["duration_number"] = pd.to_numeric(dur.str.extract(r"(\d+)")[0], errors="coerce")
        df["duration_unit"] = dur.str.extract(r"([A-Za-z]+)")[0].str.lower()
        df["duration_minutes"] = df["duration_number"].where(df["duration_unit"].str.contains("min", na=False))
        df["duration_seasons"] = df["duration_number"].where(df["duration_unit"].str.contains("season", na=False))
    else:
        if "duration_minutes" not in df.columns:
            df["duration_minutes"] = pd.NA
        if "duration_seasons" not in df.columns:
            df["duration_seasons"] = pd.NA

    if "type" in df.columns:
        df["type"] = df["type"].astype(str).str.strip()

    # Summary table
    if "type" in df.columns:
        by_type = df["type"].value_counts(dropna=False).rename_axis("type").reset_index(name="count")
        by_type.to_csv(OUT_DIR / "count_by_type.csv", index=False)

    # Chart 1: Titles by release year (by type if available)
    if "release_year" in df.columns:
        tmp = df.dropna(subset=["release_year"]).copy()
        tmp["release_year"] = pd.to_numeric(tmp["release_year"], errors="coerce")
        tmp = tmp.dropna(subset=["release_year"])
        tmp["release_year"] = tmp["release_year"].astype(int)

        if "type" in tmp.columns:
            pivot = (
                tmp.pivot_table(index="release_year", columns="type", values="title", aggfunc="count", fill_value=0)
                .sort_index()
            )
            ax = pivot.plot(figsize=(10, 5))
            ax.set_title("Netflix titles by release year (by type)")
            ax.set_xlabel("Release year")
            ax.set_ylabel("Number of titles")
            plt.tight_layout()
            plt.savefig(OUT_DIR / "titles_by_year_by_type.png", dpi=200)
            plt.close()
        else:
            counts = tmp.groupby("release_year").size().sort_index()
            plt.figure(figsize=(10, 5))
            plt.plot(counts.index, counts.values)
            plt.title("Netflix titles by release year")
            plt.xlabel("Release year")
            plt.ylabel("Number of titles")
            plt.tight_layout()
            plt.savefig(OUT_DIR / "titles_by_year.png", dpi=200)
            plt.close()

    # Chart 2: Top 10 ratings
    if "rating" in df.columns:
        ratings = df["rating"].astype(str).replace({"nan": pd.NA}).dropna()
        top = ratings.value_counts().head(10)
        plt.figure(figsize=(10, 5))
        plt.barh(top.index[::-1], top.values[::-1])
        plt.title("Top 10 Netflix ratings (by count)")
        plt.xlabel("Number of titles")
        plt.tight_layout()
        plt.savefig(OUT_DIR / "top_10_ratings.png", dpi=200)
        plt.close()

    # Chart 3: Movie duration histogram (minutes)
    mins = pd.to_numeric(df["duration_minutes"], errors="coerce").dropna()
    if len(mins) > 0:
        plt.figure(figsize=(10, 5))
        plt.hist(mins, bins=25, edgecolor="black")
        plt.title("Distribution of movie durations (minutes)")
        plt.xlabel("Duration (minutes)")
        plt.ylabel("Number of movies")
        plt.tight_layout()
        plt.savefig(OUT_DIR / "movie_duration_hist.png", dpi=200)
        plt.close()

    # Small text summary
    with open(OUT_DIR / "summary.txt", "w", encoding="utf-8") as f:
        f.write("Netflix Titles – Quick Summary\n")
        f.write("=" * 32 + "\n\n")
        f.write(f"Rows: {len(df)}\n")
        if "title" in df.columns:
            f.write(f"Unique titles: {df['title'].nunique()}\n")
        if "type" in df.columns:
            f.write(df["type"].value_counts(dropna=False).to_string())
            f.write("\n")

    print("Done! Outputs saved to:", OUT_DIR)

if __name__ == "__main__":
    main()
