-- ============================================================
-- WattWatch: Analyzing Urban Energy Consumption for Smarter Cities
-- SQL Analysis on SmartCityEnergy.csv
-- Table: SmartCityEnergy
-- Columns: MeterID, Zone, ConsumerType, Date, EnergyConsumed_kWh,
--          PeakUsage_kWh, OutageMinutes, MeterStatus, TariffRate
-- ============================================================

-- 1. Total and average daily energy consumption by zone
SELECT
    Zone,
    ROUND(SUM(EnergyConsumed_kWh), 2) AS Total_kWh,
    ROUND(AVG(EnergyConsumed_kWh), 2) AS Avg_Daily_kWh
FROM SmartCityEnergy
GROUP BY Zone
ORDER BY Total_kWh DESC;


-- 2. Top 5 highest energy-consuming consumers, by type
SELECT
    MeterID,
    ConsumerType,
    ROUND(SUM(EnergyConsumed_kWh), 2) AS Total_kWh
FROM SmartCityEnergy
GROUP BY MeterID, ConsumerType
ORDER BY Total_kWh DESC
LIMIT 5;


-- 3. Monthly trend of consumption across zones
SELECT
    Zone,
    strftime('%Y-%m', Date) AS Month,
    ROUND(SUM(EnergyConsumed_kWh), 2) AS Total_kWh
FROM SmartCityEnergy
GROUP BY Zone, Month
ORDER BY Zone, Month;


-- 4. Average cost per zone (EnergyConsumed_kWh * TariffRate)
SELECT
    Zone,
    ROUND(AVG(EnergyConsumed_kWh * TariffRate), 2) AS Avg_Cost
FROM SmartCityEnergy
GROUP BY Zone
ORDER BY Avg_Cost DESC;


-- 5. Meters with the highest number of faults / outages
SELECT
    MeterID,
    Zone,
    COUNT(*) AS Outage_Events,
    SUM(OutageMinutes) AS Total_Outage_Minutes
FROM SmartCityEnergy
WHERE OutageMinutes > 0
GROUP BY MeterID, Zone
ORDER BY Total_Outage_Minutes DESC
LIMIT 10;


-- 6. Zones with lowest energy efficiency (high usage + frequent outages)
--    Outage_per_kWh = outage minutes incurred per kWh delivered -> higher = less efficient
SELECT
    Zone,
    ROUND(SUM(EnergyConsumed_kWh), 2) AS Total_kWh,
    SUM(OutageMinutes) AS Total_Outage_Minutes,
    ROUND(SUM(OutageMinutes) * 1.0 / SUM(EnergyConsumed_kWh), 4) AS Outage_per_kWh
FROM SmartCityEnergy
GROUP BY Zone
ORDER BY Outage_per_kWh DESC;


-- 7. Peak usage patterns: weekdays vs weekends
SELECT
    CASE WHEN strftime('%w', Date) IN ('0', '6') THEN 'Weekend' ELSE 'Weekday' END AS DayType,
    ROUND(AVG(PeakUsage_kWh), 2) AS Avg_Peak_kWh
FROM SmartCityEnergy
GROUP BY DayType;
