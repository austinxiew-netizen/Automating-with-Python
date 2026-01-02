# Real Estate Data Pipeline

Python automation for consolidating and cleaning fragmented Excel market data.

### Key Logic
* **Field Mapping**: Standardizes inconsistent headers (e.g., "Vacancy %" -> "vacancy_rate").
* **Numeric Normalization**: Converts strings (e.g., "$120k", "-50k", "15%") into standard floats.
* **Row Filtering**: Automatically removes empty rows, duplicate headers, and non-data summaries.

### How to use
1. Run the script.
2. Enter the path of the folder containing `.xlsx` files.
3. Standardized results will be saved in `output_files/`.
