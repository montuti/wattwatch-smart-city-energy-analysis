# wattwatch-smart-city-energy-analysis
SQL + Excel analysis of 72K smart-meter readings across 5 city zones — demand, cost, and grid-reliability insights for smarter energy planning.
# WattWatch — Analyzing Urban Energy Consumption for Smarter Cities

An end-to-end SQL + Excel analysis of a simulated smart-city energy grid: 72,400 daily smart-meter readings across 400 meters, 5 zones, and 3 consumer segments, used to answer where energy demand, cost, and grid unreliability actually concentrate.

---
<img width="1920" height="1080" alt="Screenshot 2026-09-12 152355" src="https://github.com/user-attachments/assets/9f457045-6c9a-4c9f-a7a3-fc4f2e06b419" />

## Table of Contents

1. [Project Motivation](#project-motivation)
2. [Dataset](#dataset)
3. [Repo Structure](#repo-structure)
4. [Tools & Methodology](#tools--methodology)
5. [SQL Analysis — Query by Query](#sql-analysis--query-by-query)
6. [Findings](#findings)
7. [Data Validation](#data-validation)
8. [Dashboard](#dashboard)
9. [Business Recommendations](#business-recommendations)
10. [Limitations & Future Work](#limitations--future-work)
11. [How to Reproduce](#how-to-reproduce)
12. [Skills Demonstrated](#skills-demonstrated)
13. [Author](#author)

---

## Project Motivation

Utility companies capture meter-level data constantly, but it usually sits unused in raw logs. WattWatch simulates the analyst's job of turning that raw log into decisions a city planner or grid operator could actually act on. Three questions frame the whole project:

1. **Where is demand concentrated** — by zone and by consumer type — and is that concentration matched by cost?
2. **Where is the grid least reliable**, and is unreliability caused by a few bad meters or a systemic zone-level issue?
3. **How does demand move over time** — daily (weekday vs weekend) and monthly (seasonal)?

The project is deliberately split into two layers, mirroring a real workflow: **SQL** does the heavy aggregation/ranking, and **Excel** turns the SQL output into something a non-technical stakeholder can filter and read in 30 seconds.

## Dataset

**File:** `SmartCityEnergy.csv` — 72,400 rows, one row per meter per day.

| Property | Value |
|---|---|
| Total readings | 72,400 |
| Unique meters | 400 |
| Date range | 2025-01-01 → 2025-06-30 (181 days, 6 months) |
| Zones | North, South, East, West, Central |
| Consumer types | Residential, Commercial, Industrial |
| Meter status values | Active (69,564 rows / 96.1%), Faulty (2,836 rows / 3.9%) |
| Total energy delivered | 1,133,934.63 kWh |
| Total outage minutes logged | 340,584 |
| Total estimated cost | ₹8,314,703.45 |

### Schema

| Column | Type | Description |
|---|---|---|
| `MeterID` | text | Unique smart-meter identifier (e.g. `MTR0001`) — 400 distinct meters |
| `Zone` | text | One of the 5 city zones the meter belongs to |
| `ConsumerType` | text | `Residential`, `Commercial`, or `Industrial` |
| `Date` | date | Reading date, daily granularity, Jan–Jun 2025 |
| `EnergyConsumed_kWh` | numeric | Energy consumed that day |
| `PeakUsage_kWh` | numeric | Peak instantaneous-equivalent usage that day |
| `OutageMinutes` | integer | Minutes of outage recorded for that meter/day (0 if none) |
| `MeterStatus` | text | `Active` or `Faulty` |
| `TariffRate` | numeric | ₹ per kWh rate applied that day (ranges 5.54–9.33, mean ₹7.11) |

### Meter distribution

| Consumer Type | Unique meters | Total readings |
|---|---|---|
| Residential | 238 | 43,078 |
| Commercial | 123 | 22,263 |
| Industrial | 39 | 7,059 |

| Zone | Unique meters | Active readings | Faulty readings |
|---|---|---|---|
| East | 88 | 15,440 | 488 |
| West | 87 | 15,036 | 711 |
| North | 82 | 14,213 | 629 |
| South | 73 | 12,748 | 465 |
| Central | 70 | 12,127 | 543 |

West has the highest faulty-reading count in absolute terms (711), but this needs to be read against total readings per zone (see Findings #3 for the outage-rate-normalized view, which tells a different story than raw faulty-row counts).
<img width="817" height="546" alt="Screenshot 2026-09-12 152019" src="https://github.com/user-attachments/assets/aadbc5fa-18f9-4966-8fd0-a0ccfbb1048f" />

## Repo Structure

```
wattwatch/
├── SmartCityEnergy.csv        # raw dataset (72,400 rows)
├── SmartCity_SQL.sql          # all analysis queries (7 queries)
├── WattWatchDashboard.xlsx    # Excel dashboard: Dashboard / RawData / Summary sheets
├── EnergyFindings.md          # written findings (corrected version — see Findings below)
├── assets/
│   └── dashboard.png          # dashboard screenshot (add this — see Dashboard section)
└── README.md
```

## Tools & Methodology

- **SQL (SQLite dialect)** — used for all aggregation, ranking, and derived-metric calculation. SQLite's `strftime()` is used for date-part extraction (month, day-of-week), which is why the queries won't run unmodified on MySQL/Postgres (see Reproduce section for the fix).
- **Excel** — a 3-sheet workbook (`RawData`, `Summary`, `Dashboard`). `RawData` holds the CSV as a table with AutoFilter headers acting as slicers; `Summary` holds the aggregation formulas backing the KPI cards; `Dashboard` is the presentation layer with 4 linked charts and 4 KPI cards.
- **Python/Pandas** — not part of the deliverable, but used as an independent second implementation to verify every number in this README against the raw CSV before publishing (see Data Validation).

The general method: write a SQL query → sanity-check the row count/shape → pull the result into the Excel workbook or restate it as a finding → re-derive the same number independently in Pandas → only keep the finding if both methods agree.

## SQL Analysis — Query by Query

### 1. Total & average daily energy consumption by zone
```sql
SELECT Zone,
       ROUND(SUM(EnergyConsumed_kWh), 2) AS Total_kWh,
       ROUND(AVG(EnergyConsumed_kWh), 2) AS Avg_Daily_kWh
FROM SmartCityEnergy
GROUP BY Zone
ORDER BY Total_kWh DESC;
```
Ranks zones by total demand over the 6 months. This is the starting point — everything else (cost, reliability, seasonality) gets read in the context of who's actually using the most power.

### 2. Top 5 highest energy-consuming meters, by type
```sql
SELECT MeterID, ConsumerType,
       ROUND(SUM(EnergyConsumed_kWh), 2) AS Total_kWh
FROM SmartCityEnergy
GROUP BY MeterID, ConsumerType
ORDER BY Total_kWh DESC
LIMIT 5;
```
Individual-meter view, not zone/type-level — surfaces the single heaviest loads in the whole grid, useful for capacity planning and infrastructure prioritization at specific sites.

### 3. Monthly trend of consumption across zones
```sql
SELECT Zone, strftime('%Y-%m', Date) AS Month,
       ROUND(SUM(EnergyConsumed_kWh), 2) AS Total_kWh
FROM SmartCityEnergy
GROUP BY Zone, Month
ORDER BY Zone, Month;
```
Builds the time series behind the "Monthly kWh Trend by Zone" dashboard chart — this is what exposes the spring seasonal bump (see Findings #7).

### 4. Average cost per zone
```sql
SELECT Zone,
       ROUND(AVG(EnergyConsumed_kWh * TariffRate), 2) AS Avg_Cost
FROM SmartCityEnergy
GROUP BY Zone
ORDER BY Avg_Cost DESC;
```
Cost and consumption don't necessarily rank the same way, because `TariffRate` varies by reading (5.54–9.33). This query is what surfaces Central as the highest-cost zone despite not having the highest total kWh (Finding #5).

### 5. Meters with the highest number of faults/outages
```sql
SELECT MeterID, Zone,
       COUNT(*) AS Outage_Events,
       SUM(OutageMinutes) AS Total_Outage_Minutes
FROM SmartCityEnergy
WHERE OutageMinutes > 0
GROUP BY MeterID, Zone
ORDER BY Total_Outage_Minutes DESC
LIMIT 10;
```
Filters to only rows with a recorded outage, then ranks individual meters — this is a hardware-inspection worklist, not a zone-level metric.

### 6. Zones with lowest energy efficiency
```sql
SELECT Zone,
       ROUND(SUM(EnergyConsumed_kWh), 2) AS Total_kWh,
       SUM(OutageMinutes) AS Total_Outage_Minutes,
       ROUND(SUM(OutageMinutes) * 1.0 / SUM(EnergyConsumed_kWh), 4) AS Outage_per_kWh
FROM SmartCityEnergy
GROUP BY Zone
ORDER BY Outage_per_kWh DESC;
```
This is the most important query in the set: it normalizes outage minutes by kWh delivered, so a zone doesn't look "reliable" just because it happens to deliver more power. Without this normalization, North (highest total outage-adjacent readings in absolute terms) would look worse than it actually is relative to how much it delivers.

### 7. Peak usage: weekdays vs weekends
```sql
SELECT CASE WHEN strftime('%w', Date) IN ('0','6') THEN 'Weekend' ELSE 'Weekday' END AS DayType,
       ROUND(AVG(PeakUsage_kWh), 2) AS Avg_Peak_kWh
FROM SmartCityEnergy
GROUP BY DayType;
```
`strftime('%w', Date)` returns `0` for Sunday and `6` for Saturday in SQLite. Everything else buckets as weekday.
<img width="1346" height="495" alt="Screenshot 2026-09-12 152056" src="https://github.com/user-attachments/assets/201cd166-18d0-413e-8db4-a7238c51223e" />

## Findings

**1. North and Central are the city's demand hotspots.** North totals 260,776 kWh over 6 months, Central 245,377 kWh — both meaningfully above South (184,304 kWh, the lowest of the five zones). East (222,803 kWh) and West (220,674 kWh) sit in between. North also has the most unique meters after East (82 vs 88), so its lead isn't purely a meter-count artifact — its per-meter average is also higher.

**2. Residential is the largest consumer segment by total volume — not Industrial.** This corrects an earlier draft of the findings.
   - Residential: 448,921 kWh total (238 meters, 43,078 readings)
   - Commercial: 411,188 kWh total (123 meters, 22,263 readings)
   - Industrial: 273,826 kWh total (39 meters, 7,059 readings) — **lowest** of the three, despite each meter running hardest
   - The confusion comes from a real but separate fact: the **top 5 single highest-consuming meters in the entire dataset are all Industrial** (8,877–9,020 kWh each over 6 months, roughly 4–5× a typical residential meter). Industrial meters are individually the heaviest loads, but there are so few of them (39, ~10% of all meters) that their category total ends up smallest. Per-meter intensity and category-level total are two different questions, and this dataset answers them oppositely.

**3. South and East are the least reliable zones, normalized for delivery.** Outage-minutes-per-kWh:

   | Zone | Total kWh | Total Outage Min | Outage/kWh |
   |---|---|---|---|
   | South | 184,304 | 72,807 | **0.3950** |
   | East | 222,803 | 77,896 | **0.3496** |
   | West | 220,674 | 67,372 | 0.3053 |
   | North | 260,776 | 68,677 | 0.2634 |
   | Central | 245,377 | 53,832 | 0.2194 |

   Note this ranking is *not* the same as ranking by raw faulty-reading count (West had the most faulty rows, 711) — South's smaller total delivery makes its outage burden look proportionally worse even though its raw outage-minute total (72,807) isn't the highest either (East's is, at 77,896). Grid-maintenance priority should follow this normalized ranking (South → East → West → North → Central), not the raw counts.

**4. A small set of meters is responsible for most reliability problems.** The 10 worst-performing meters (concentrated in South, Central, North) each logged 13–25 separate outage events totaling 2,244–3,239 minutes. The single worst is `MTR0281` (South): 25 events, 3,239 minutes. These 10 meters are prime candidates for physical hardware inspection or replacement rather than zone-wide infrastructure spend.

**5. Central has the highest average cost per reading (₹140.37)** despite ranking 2nd, not 1st, in total consumption. This is a tariff-mix effect, not a volume effect — meaning cost-reduction interventions (e.g. renegotiating tariff structure, shifting load to off-peak) may have a larger ₹-impact in Central than in North, even though North uses more total power.

**6. Peak demand is a weekday phenomenon.** Average peak usage: weekday 4.18 kWh vs weekend 3.28 kWh — a 27.5% gap, consistent with commercial/industrial activity driving the city's load curve. This also means demand-response programs (asking large users to shift load) would have the most room to work on weekdays.

**7. Consumption follows a mild seasonal curve**, dipping in February and rising through March–May before falling back in June:

   | Month | Central | East | North | South | West |
   |---|---|---|---|---|---|
   | Jan | 40,156 | 36,457 | 42,681 | 30,103 | 36,135 |
   | Feb | 38,347 | 34,699 | 40,670 | 28,738 | 34,437 |
   | Mar | 43,349 | 39,294 | 46,101 | 32,634 | 38,980 |
   | Apr | 42,484 | 38,802 | 45,469 | 32,041 | 38,258 |
   | May | 42,488 | 38,477 | 45,052 | 31,875 | 38,103 |
   | Jun | 38,553 | 35,074 | 40,804 | 28,913 | 34,761 |

   All five zones move together — the dip and rise happen in the same months everywhere, which suggests a shared external driver (most likely temperature/HVAC load) rather than a zone-specific cause. Worth correlating against actual weather data if that becomes available.

## Data Validation

Every number in this README was independently re-derived in Pandas directly from `SmartCityEnergy.csv`, rather than copied from the SQL output or an earlier findings draft. The overall totals also cross-check against the Excel dashboard's KPI cards:

| Metric | CSV/Pandas | Excel Dashboard KPI |
|---|---|---|
| Total consumption | 1,133,934.63 kWh | 1,133,934.63 |
| Avg peak usage | 3.9198 kWh | 3.91976326 |
| Total outage minutes | 340,584 | 340,584 |
| Total estimated cost | ₹8,314,703.45 | ₹8,314,703.451 |

All four match, which confirms the SQL logic, the CSV, and the Excel formulas are internally consistent. The one figure that did **not** survive validation was the "Industrial dominates total usage" claim from an earlier findings draft — replaced with the verified breakdown in Finding #2.

## Dashboard

`WattWatchDashboard.xlsx` — 3 sheets:

- **RawData** — the full 72,400-row CSV as an Excel Table, with AutoFilter dropdowns on `Zone`, `ConsumerType`, and `Date` in the header row. These act as the report's slicers — filtering here should be reflected in `Summary` and `Dashboard` if the formulas are range/table-referenced correctly.
- **Summary** — aggregation formulas backing the KPI cards and charts.
- **Dashboard** — the presentation layer:
  - 4 KPI cards: Total Consumption (1,133,934.63 kWh), Avg Peak Usage (3.92 kWh), Total Outage Minutes (340,584), Total Estimated Cost (₹8,314,703.45)
  - Energy Usage by Consumer Type (bar)
  - Energy Use by Zone (bar)
  - Weekday vs Weekend Avg Peak Usage (bar)
  - Monthly kWh Trend by Zone (line, 5 series)

**To add:** export the Dashboard sheet as an image and save it to `assets/dashboard.png`, then add:
```markdown
![WattWatch Dashboard](assets/dashboard.png)
```
right here in the README so it renders on GitHub without opening the workbook.

## Business Recommendations

1. **Prioritize grid maintenance in South, then East** — not by raw outage-minute totals, but by the normalized outage-per-kWh ratio, which is the metric that actually reflects reliability relative to what's being delivered.
2. **Send the 10 worst meters (led by MTR0281) for physical inspection/replacement** before any broader zone-level infrastructure spend — the outage problem is concentrated in specific hardware, not distributed evenly.
3. **Target demand-response or load-shifting programs at weekdays**, since that's where the 27.5% peak-usage gap lives — weekend infrastructure is comparatively under-stressed.
4. **Investigate Central's tariff mix** before assuming its high average cost (₹140.37) is a usage problem — it isn't the highest-consumption zone, so the lever there is pricing/tariff structure, not demand reduction.
5. **Don't design Industrial-specific efficiency programs assuming it's the biggest opportunity by volume** — Residential is. Industrial is the right target if the goal is *peak load reduction from a few sites*; Residential is the right target if the goal is *total citywide kWh reduction*.

## Limitations & Future Work

- **No weather data.** The seasonal dip/rise (Finding #7) is almost certainly temperature-driven, but this can't be confirmed without external temperature data for the same period.
- **`OutageMinutes` and `MeterStatus` are point-in-time daily fields**, not full outage-event logs — the "outage events" count in Query 5 counts *days with any outage*, not necessarily distinct physical incidents (a single multi-day outage would count as multiple "events" here).
- **No customer count or floor-area normalization.** Comparing zones on raw kWh assumes the population/building density per zone is comparable, which isn't verified in this dataset.
- **SQLite-specific date functions.** `strftime()` in Queries 3 and 7 won't run as-is on MySQL/Postgres — see Reproduce below for the fix.
- **Future iteration ideas:** correlate monthly trend against real temperature data; build an anomaly-detection pass on `OutageMinutes` to distinguish single long outages from repeated short ones; add a per-capita or per-meter-density normalized zone comparison.

## How to Reproduce

1. Load `SmartCityEnergy.csv` into a SQLite database as table `SmartCityEnergy` (or use the `sqlite3` CLI: `.import --csv SmartCityEnergy.csv SmartCityEnergy`).
2. Run `SmartCity_SQL.sql` against it.
3. **Porting to MySQL/Postgres:** replace `strftime('%Y-%m', Date)` with `DATE_FORMAT(Date, '%Y-%m')` (MySQL) or `TO_CHAR(Date, 'YYYY-MM')` (Postgres); replace `strftime('%w', Date)` with `DAYOFWEEK(Date)` (MySQL, 1=Sunday) or `EXTRACT(DOW FROM Date)` (Postgres, 0=Sunday) and adjust the weekend check accordingly.
4. Open `WattWatchDashboard.xlsx`, use the `RawData` sheet's AutoFilter dropdowns to slice by Zone/ConsumerType/Date, and confirm the Dashboard charts update.

## Skills Demonstrated

- **SQL:** `GROUP BY` with multi-column grouping, aggregate functions (`SUM`, `AVG`, `COUNT`), `ORDER BY`/`LIMIT` ranking, `CASE WHEN` bucketing, date-part extraction, derived-ratio metrics computed inline
- **Excel:** AutoFilter-based slicers, multi-chart dashboard layout with linked KPI cards, table-referenced formulas
- **Analytical rigor:** cross-validating every written finding against the raw dataset in a second tool (Pandas) before publishing, and correcting a finding that didn't hold up rather than leaving it in

## Author

**Montu** — Data Analysis student, RW Skill Education (Sirohi, Rajasthan)
GitHub: [github.com/montuti](https://github.com/montuti)
sql excel data-analysis data-analytics energy-analytics dashboard sqlite smart-city
