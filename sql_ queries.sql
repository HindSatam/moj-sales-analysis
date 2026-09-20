-- Assumptions that apply to both queries:
--  * The queries run against the output of clean_data.py (cleansed_data.csv) loaded as
--    dbo.moj_sales_cleansed. It has the columns in schema.md plus three derived ones:
--    deed_date (Gregorian), [quarter] (text such as '2020Q1') and [value] (AREA * METER_PRICE, in SAR).
--  * In the cleansed table missing text values are NULL, and rows with a missing/zero price or area
--    and extreme outliers have already been removed.
--  * "Quarter" means a Gregorian calendar quarter, and one row is treated as one transaction
--    (the schema has no deed identifier, and only exact duplicate rows were removed).

DECLARE @Quarter char(6) = '2020Q1';


-- For a given quarter, return the top 10 districts by average sale price for
-- residential properties, excluding any district with fewer than 30 transactions.

SELECT TOP (10)
    NEIGHBORHOOD_NAME AS district,
    AVG([value])      AS avg_sale_price,
    COUNT(*)          AS transaction_count
FROM dbo.moj_sales_cleansed
WHERE [quarter] = @Quarter
  AND new_use = N'سكني'
  AND NEIGHBORHOOD_NAME IS NOT NULL
GROUP BY NEIGHBORHOOD_NAME
HAVING COUNT(*) >= 30
ORDER BY avg_sale_price DESC, NEIGHBORHOOD_NAME;

-- What it does: ranks districts by the average transaction value (area x price per m2) of their
-- residential sales in the quarter and returns the 10 highest, keeping only districts with at least 30 sales.
-- Assumptions: "sale price" is the full transaction value, not the price per m2 (METER_PRICE is a unit price);
-- residential is defined by new_use rather than RS_TYPE, which holds land and apartment descriptions;
-- rows with no district are left out; ties in the average are broken by district name.


-- For the same quarter, return per city: the number of transactions,
-- the number with a missing district, and the number with a missing or zero price.

SELECT
    RS_CITY_NAME AS city,
    COUNT(*)     AS transaction_count,
    SUM(CASE WHEN NEIGHBORHOOD_NAME IS NULL
             THEN 1 ELSE 0 END) AS missing_district_count,
    SUM(CASE WHEN METER_PRICE IS NULL OR METER_PRICE = 0
             THEN 1 ELSE 0 END) AS missing_or_zero_price_count
FROM dbo.moj_sales_cleansed
WHERE [quarter] = @Quarter
GROUP BY RS_CITY_NAME
ORDER BY transaction_count DESC;

-- What it does: counts each city's transactions in the quarter, and how many of them have no district
-- and how many have a missing or zero price.
-- Assumptions: all property types are included (no residential filter, since the request only says "the same quarter");
-- on the cleansed table the price count is always 0 because clean_data.py already removes those rows, so it only
-- shows a non-zero result on the raw extract; rows with no city appear as one group with a NULL city.
