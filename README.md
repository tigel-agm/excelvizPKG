# ExcelVizLIB

ExcelViz is a Python library designed to streamline the process of loading, cleaning, visualizing, and summarizing data directly from Excel files (`.xlsx`, `.xls`). It provides a convenient wrapper around popular libraries like Pandas, Matplotlib, Seaborn, NLTK, and others for common data analysis workflows.

![diagram](https://github.com/user-attachments/assets/b835667c-6d16-43d1-9cf2-3e023fdd291f)

## Features

*   **Flexible Excel Loading:** Load data from specific sheets, rows, and columns using various engines (`openpyxl`, `calamine`). Includes option to read formulas as strings (requires `openpyxl`).
*   **Data Cleaning:** Handle missing values (various strategies including ffill/bfill), standardize text using fuzzy matching, and clean text for analysis (stopwords, stemming).
*   **Data Manipulation:** Reshape data (`pivot_data`, `melt_data`), discretize (`bin_numeric_column`), create dummies (`create_dummy_variables`), merge (`merge_dataframes`), aggregate (`group_and_aggregate`), feature engineering (`create_interaction_term`, `extract_datetime_components`).
*   **Outlier Detection:** Identify outliers using the IQR method (`detect_outliers_iqr`). (More methods can be added).
*   **Data Profiling:** Generate a basic profiling report with stats and histograms (`generate_profile_report`).
*   **Visualization:** Generate common plots directly from DataFrames:
    *   Bar charts (including count plots, grouped, stacked)
    *   Line plots
    *   Scatter plots (static, with save format options)
    *   Treemaps
    *   Pie charts
    *   Word clouds (with integrated text cleaning)
    *   Correlation heatmaps
    *   Histograms (with KDE)
    *   Box plots
    *   Interactive Scatter Plots (HTML output via Plotly) (More interactive types can be added)
*   **Statistical Summaries:** Generate descriptive statistics, perform basic multivariate analysis (PCA, MCA, etc.) using `prince`. (Advanced stats via `statsmodels` can be added).
*   **Excel I/O:** Read formulas (`load_excel`), save multiple DataFrames to different sheets with basic formatting (`save_multi_sheet_excel`).
*   **Formula Evaluation (Experimental):** Attempt to evaluate Excel formulas using the `formulas` library (`evaluate_formulas`). *Use with caution!*
*   **Data Export:** Save processed DataFrames back to Excel files.

## Installation

1.  **Clone or Download:** Get the `excelviz` package directory.
2.  **Install Dependencies:** Ensure you have Python 3.x installed. Then, install the required libraries. It's recommended to use a virtual environment.

    ```bash
    pip install pandas numpy openpyxl matplotlib seaborn plotnine squarify wordcloud nltk snowballstemmer statsmodels prince Pillow thefuzz[speedup] tabulate calamine Pillow
    ```
    *(Note: `plotnine` is listed as a dependency but wasn't explicitly used in the final functions; it can be omitted if not needed for future extensions)*

3.  **Download NLTK Data:** `ExcelViz` uses NLTK for text cleaning (stopwords, tokenization). Download the necessary resources by running this command in your terminal *once*:

    ```bash
    python -m nltk.downloader punkt stopwords wordnet punkt_tab
    ```

4.  **Place `excelviz`:** Ensure the `excelviz` package directory is in your Python path or in the same directory as your script that imports it.

## Basic Usage

```python
import pandas as pd
import excelviz
import os

# --- Configuration ---
DATA_FILE = 'your_data.xlsx' # Path to your Excel file
OUTPUT_DIR = 'analysis_outputs'
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- 1. Load Data ---
try:
    df = excelviz.load_excel(DATA_FILE, sheet_name='Sheet1')
    print("Data loaded successfully.")
except FileNotFoundError:
    print(f"Error: File not found at {DATA_FILE}")
    exit()
except Exception as e:
    print(f"Error loading data: {e}")
    exit()

# --- 2. Clean Data (Example: Fill missing numeric values) ---
# Assuming 'NumericColumn' might have missing values
if 'NumericColumn' in df.columns:
    df_cleaned = excelviz.handle_missing_values(
        df,
        strategy='fill_median',
        subset=['NumericColumn']
    )
    print("Missing values handled.")
else:
    df_cleaned = df.copy() # Or handle other columns

# --- 3. Visualize (Example: Bar chart of a 'Category' column) ---
if 'Category' in df_cleaned.columns:
    try:
        excelviz.plot_bar(
            df_cleaned,
            x_col='Category',
            title='Distribution of Categories',
            output_path=os.path.join(OUTPUT_DIR, 'category_distribution.png')
        )
    except Exception as e:
        print(f"Error creating bar chart: {e}")

# --- 4. Summarize (Example: Descriptive statistics) ---
try:
    summary_string = excelviz.generate_descriptive_summary(
        df_cleaned,
        format_output=True # Get a formatted string
    )
    print("\n--- Descriptive Summary ---")
    print(summary_string)
except Exception as e:
    print(f"Error generating summary: {e}")

# --- 5. Save Cleaned Data (Optional) ---
try:
    excelviz.save_to_excel(
        df_cleaned,
        os.path.join(OUTPUT_DIR, 'cleaned_output.xlsx'),
        sheet_name='CleanedData'
    )
except Exception as e:
    print(f"Error saving cleaned data: {e}")

print(f"\nAnalysis complete. Outputs saved in '{OUTPUT_DIR}'.")

```

## API Overview

*(See also `demo.py` for basic usage and `advanced_demo.py` for examples of more features. Use `cli.py` to run demos and data generation from the command line.)*

*(Refer to the docstrings within each function in the `excelviz` modules for detailed parameter descriptions.)*

*   **Loading:**
    *   `excelviz.load_excel(file_path, sheet_name=0, header=0, usecols=None, engine='openpyxl', read_formulas=False)`
    *   `excelviz.save_multi_sheet_excel(dataframes, output_path, formats=None, bold_header=True, ...)`
*   **Processing & Manipulation:**
    *   `excelviz.handle_missing_values(df, strategy='drop_row', subset=None, fill_value=None, threshold=None)` (Strategies: 'drop_row', 'drop_col', 'fill_mean', 'fill_median', 'fill_mode', 'fill_constant', 'ffill', 'bfill')
    *   `excelviz.clean_text_column(df, column_name, lowercase=True, remove_punctuation=True, remove_stopwords=True, language='english', stemming=True, stemmer_language=None)`
    *   `excelviz.fuzzy_match_standardize(df, column_name, reference_list, score_cutoff=80)`
    *   `excelviz.save_to_excel(df, output_path, sheet_name='Sheet1', **to_excel_kwargs)`
    *   `excelviz.pivot_data(df, index_cols, column_cols, value_cols, aggfunc='mean', fill_value=None)`
    *   `excelviz.melt_data(df, id_vars, value_vars=None, var_name=None, value_name='value')`
    *   `excelviz.bin_numeric_column(df, column_name, bins, labels=None, new_column_name=None)`
    *   `excelviz.create_dummy_variables(df, columns, prefix=None, drop_first=False)`
    *   `excelviz.merge_dataframes(left_df, right_df, how='inner', on=None, ...)`
    *   `excelviz.group_and_aggregate(df, group_by_cols, agg_dict, as_index=False)`
    *   `excelviz.create_interaction_term(df, col1, col2, new_column_name=None)`
    *   `excelviz.extract_datetime_components(df, column_name, components=[...])`
*   **Formula Evaluation (Experimental):**
    *   `excelviz.evaluate_formulas(excel_path, sheet_name=None, output_format='dataframe')`
*   **Outlier Detection:**
    *   `excelviz.detect_outliers_iqr(df, column_name, threshold=1.5)`
*   **Data Profiling:**
    *   `excelviz.generate_profile_report(df, title="Data Profile Report", output_dir="profile_report", ...)`
*   **Visualization:**
    *   `excelviz.plot_bar(df, x_col, y_col=None, hue_col=None, title="Bar Chart", ..., stacked=False, **kwargs)`
    *   `excelviz.plot_line(df, x_col, y_col, hue_col=None, title="Line Plot", ..., output_path="line_plot.png", **kwargs)`
    *   `excelviz.plot_scatter(df, x_col, y_col, hue_col=None, size_col=None, style_col=None, title="Scatter Plot", ..., output_path="scatter_plot.png", save_format=None, **kwargs)`
    *   `excelviz.plot_treemap(df, size_col, label_col, color_col=None, title="Treemap", ..., output_path="treemap.png", **squarify_kwargs)`
    *   `excelviz.plot_pie(df, category_col, title="Pie Chart", ..., output_path="pie_chart.png", **pie_kwargs)`
    *   `excelviz.plot_wordcloud(df, text_col, output_path="wordcloud.png", use_cleaned_text=True, clean_text_args=None, **wordcloud_kwargs)`
    *   `excelviz.plot_correlation_heatmap(df, title="Correlation Heatmap", ..., output_path="correlation_heatmap.png", **heatmap_kwargs)`
    *   `excelviz.plot_histogram(df, column_name, title=None, ..., output_path="histogram.png", bins='auto', kde=True, **hist_kwargs)`
    *   `excelviz.plot_boxplot(df, x_col=None, y_col=None, hue_col=None, title=None, ..., output_path="boxplot.png", **box_kwargs)`
    *   `excelviz.plot_interactive_scatter(df, x_col, y_col, color_col=None, size_col=None, ..., output_path="interactive_scatter.html", **scatter_px_kwargs)`
*   **Summarization:**
    *   `excelviz.generate_descriptive_summary(df, include_dtypes=True, format_output=True, tablefmt='pretty', columns=None)`
    *   `excelviz.generate_multivariate_summary(df, analysis_type='pca', n_components=2, columns=None, **kwargs)`

## License
Private Lib for now until I have worked out the bugs and edge cases

## Command-Line Interface

A basic CLI (`cli.py`) is provided to run the demo scripts:

*   Generate advanced sample data: `python cli.py generate_adv_data`
*   Run the basic demo: `python cli.py run_demo`
*   Run the advanced demo: `python cli.py run_adv_demo` (will generate data if needed)
*   Run all generation and demos: `python cli.py run_all`

## Testing

Basic unit tests are located in the `tests/` directory and can be run using `pytest` (install with `pip install pytest`). More comprehensive tests are needed.

```bash
pytest tests/
```

## Sample Outputs
<img src="https://github.com/user-attachments/assets/519db37e-167a-4b94-b393-ebb54d04c633" alt="category_barplot" width="400">
<img src="https://github.com/user-attachments/assets/93926bcd-4a4d-44ab-b667-6f0baba53f5e" alt="correlation_heatmap" width="400">
<img src="https://github.com/user-attachments/assets/2b9b1735-2f66-4665-b24d-07ff9288e77b" alt="feedback_wordcloud" width="400">
<img src="https://github.com/user-attachments/assets/8b992338-1491-47e9-b6d0-c2ed51a102cf" alt="profit_sales_scatter" width="400">
<img src="https://github.com/user-attachments/assets/9a66d9e9-783c-4c70-8c5d-a8cfe2bfcecc" alt="region_pie" width="400">
<img src="https://github.com/user-attachments/assets/b5c5b024-a7c5-4ed7-aa63-5d17215bdafa" alt="sales_over_time_lineplot" width="400">
<img src="https://github.com/user-attachments/assets/ed41fc6f-30e5-4ddb-8617-197878eadbec" alt="sales_region_category_barplot" width="400">
<img src="https://github.com/user-attachments/assets/a022eda9-65bc-4e88-9e2d-9616ebf592f6" alt="units_category_treemap" width="400">






