# California Housing Data Analysis

A data-cleaning and exploratory analysis project performed in Excel using formula-driven calculations (no hardcoded values) on the California Housing dataset (20,640 records).

## Dataset

| Column | Description |
|---|---|
| `longitude`, `latitude` | Geographic coordinates of the housing block |
| `housing_median_age` | Median age of houses in the block |
| `total_rooms`, `total_bedrooms` | Total rooms/bedrooms in the block |
| `population`, `households` | Population and household counts |
| `median_income` | Median income of households (in tens of thousands USD) |
| `median_house_value` | Median house value (target variable) |
| `ocean_proximity` | Categorical location type: `NEAR BAY`, `<1H OCEAN`, `INLAND`, `NEAR OCEAN`, `ISLAND` |

**Source rows:** 20,640
**File:** `housing.csv` → analyzed in `housing_analysis.xlsx`

## Objective

Practice core data-analyst skills — aggregation, filtering, grouping, data cleaning, and outlier detection — using only Excel formulas (fully auditable, no manual overrides).

## Data Cleaning

### Missing value check
Used `COUNTBLANK()` across every column to check for gaps.

```excel
=COUNTBLANK(range)
```

**Finding:** All columns are complete except `total_bedrooms`, which has **207 missing values** (~1% of records).

### Handling missing values
Two new columns were engineered rather than deleting or ignoring the incomplete rows:

**`total_bedrooms_filled`** — imputes missing values with the column median:
```excel
=IF(E2="",MEDIAN($E$2:$E$20641),E2)
```
*Logic:* if the original cell is blank, substitute the median of the full column; otherwise keep the original value untouched.

**`is_bedrooms_missing`** — flags which rows were imputed, for transparency:
```excel
=IF(E2="","YES","NO")
```
*Logic:* this creates an audit trail so imputed values are never mistaken for real observations.

## Analysis Performed

### 1. Aggregate functions
- Average `median_house_value` per `ocean_proximity` category — `AVERAGEIF()`
- Max/min `median_income` — `MAX()`, `MIN()`
- Record count per category — `COUNTIF()`
- Total population — `SUM()`
- Average housing age for high-value homes (>$300,000) — `AVERAGEIF()`

### 2. Filtering
- Multi-condition filtering (value + location) — `COUNTIFS()`
- Ratio-based filtering (bedrooms per household) — `SUMPRODUCT()`
- Missing-data row detection — `SUMPRODUCT()`

### 3. Grouping
- Category-level averages compared against a threshold — `AVERAGEIF()` + manual comparison
- Highest-value location category — `AVERAGEIF()` + `MAX()`

### 4. Derived metrics
- `rooms_per_household` = `total_rooms / households`
- `bedroom_ratio` = `total_bedrooms / total_rooms`
- Top 10 highest-value properties — `LARGE()`

### 5. Outlier detection
Applied the 2-standard-deviation rule on `median_income`:
```excel
Upper bound = AVERAGE(income) + 2 * STDEVP(income)
Lower bound = AVERAGE(income) - 2 * STDEVP(income)
```
Values outside this range are flagged as statistical outliers.

## Key Findings

- **Data quality:** 99% complete; only `total_bedrooms` required cleaning (207 rows, handled via median imputation with a transparency flag).
- **Location matters:** Average `median_house_value` varies substantially by `ocean_proximity` — coastal/bay categories command a premium over `INLAND`.
- **Income is the strongest visible driver:** Higher `median_income` blocks correspond to higher `median_house_value`, consistent with expected real-estate pricing behavior.

## Visualization

A bar chart was added to the `Analysis` sheet: **"Average Median House Value by Ocean Proximity."**
- X-axis: `ocean_proximity` category
- Y-axis: average `median_house_value`
- Built directly from the Q1 aggregate formulas, so it updates automatically if the underlying data changes.

**Insight:** `ISLAND` and `NEAR BAY` properties show the highest average value, while `INLAND` properties are clearly the lowest — visually confirming that proximity to the ocean is associated with higher housing prices.

## Tools Used

- Microsoft Excel (formula-based, zero hardcoded results — fully recalculates if source data changes)
- Functions: `AVERAGEIF`, `AVERAGEIFS`, `COUNTIF`, `COUNTIFS`, `SUMPRODUCT`, `MEDIAN`, `MAX`, `MIN`, `LARGE`, `STDEVP`, `IF`, `COUNTBLANK`

## Next Steps

- Add correlation analysis between `median_income` and `median_house_value`
- Add a scatter plot for income vs. house value
- Rebuild the workflow in Python (pandas) for portfolio diversification

---
*Author: Khushbakhat | BS Information Technology, GCUF |Junior Data Analyst
