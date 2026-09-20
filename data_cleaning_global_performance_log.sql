-- ============================================================================
-- GLOBAL PERFORMANCE DAILY LOG — DATA CLEANING SCRIPT (MySQL 8.0+)
-- Author: Syafriandi | Sample documentation, sanitized for portfolio use
--
-- Source: Global_Performance_Daily_Log_RAW.xlsx (Idle_Log, AHT_Log sheets)
-- Purpose: Clean 13 known issue types (see Data Cleaning & Preparation doc)
--          before connecting the output to Tableau.
--
-- HOW TO USE:
--   1. Export each Excel sheet (Idle_Log, AHT_Log, Team Member List) to CSV.
--   2. Import each CSV into the staging_* tables below via MySQL Workbench's
--      Table Data Import Wizard — load every column as TEXT, do NOT let the
--      wizard auto-detect types (mixed date formats / percent-strings will
--      silently mis-import otherwise).
--   3. Run this script top to bottom.
--   4. Connect Tableau to clean_idle_log and clean_aht_log at the end.
-- ============================================================================


-- ----------------------------------------------------------------------------
-- 0. STAGING TABLES — raw import target, everything as text/loose types
-- ----------------------------------------------------------------------------

DROP TABLE IF EXISTS staging_idle_log;
CREATE TABLE staging_idle_log (
    row_id            INT AUTO_INCREMENT PRIMARY KEY,
    raw_date          VARCHAR(30),
    week              VARCHAR(10),
    region            VARCHAR(30),
    country           VARCHAR(10),
    tl                VARCHAR(30),
    name              VARCHAR(60),
    audit_stage       VARCHAR(30),
    product_level_1   VARCHAR(30),
    product_level_2   VARCHAR(40),
    idle_task_rate    VARCHAR(20),   -- text on purpose: mixes "0.095" and "9.5%"
    avg_idle_duration DECIMAL(10,2),
    max_idle_duration DECIMAL(10,2)
);

DROP TABLE IF EXISTS staging_aht_log;
CREATE TABLE staging_aht_log (
    row_id            INT AUTO_INCREMENT PRIMARY KEY,
    raw_date          VARCHAR(30),
    week              VARCHAR(10),
    region            VARCHAR(30),
    country           VARCHAR(10),
    tl                VARCHAR(30),
    name              VARCHAR(60),
    audit_stage       VARCHAR(30),
    product_level_1   VARCHAR(30),
    product_level_2   VARCHAR(40),
    aht_sec           VARCHAR(20),   -- text on purpose: mixes numbers and "N/A"/"TBD"
    aht_target_sec    DECIMAL(10,2),
    meets_target      VARCHAR(20)
);

-- Reference table used to back-fill missing TL / Region from Name
DROP TABLE IF EXISTS team_member_list;
CREATE TABLE team_member_list (
    name    VARCHAR(60) PRIMARY KEY,
    region  VARCHAR(30),
    market  VARCHAR(60),
    tl      VARCHAR(30)
);

-- Import your CSVs into the three tables above before continuing.


-- ----------------------------------------------------------------------------
-- 1. REMOVE FULLY BLANK ROWS   (issue: "Fully blank rows", 9 rows)
-- ----------------------------------------------------------------------------

DELETE FROM staging_idle_log
WHERE raw_date IS NULL AND name IS NULL AND idle_task_rate IS NULL;

DELETE FROM staging_aht_log
WHERE raw_date IS NULL AND name IS NULL AND aht_sec IS NULL;


-- ----------------------------------------------------------------------------
-- 2. REMOVE EXACT DUPLICATE ROWS   (issue: "Duplicate rows", 44 rows)
--    Keeps the first occurrence (lowest row_id) of each duplicate group.
--    <=> is MySQL's NULL-safe equals, so NULLs match NULLs correctly.
-- ----------------------------------------------------------------------------

DELETE t1 FROM staging_idle_log t1
JOIN staging_idle_log t2
  ON  t1.row_id > t2.row_id
  AND t1.raw_date          <=> t2.raw_date
  AND t1.name               <=> t2.name
  AND t1.tl                 <=> t2.tl
  AND t1.idle_task_rate     <=> t2.idle_task_rate
  AND t1.avg_idle_duration  <=> t2.avg_idle_duration;

DELETE t1 FROM staging_aht_log t1
JOIN staging_aht_log t2
  ON  t1.row_id > t2.row_id
  AND t1.raw_date  <=> t2.raw_date
  AND t1.name      <=> t2.name
  AND t1.tl        <=> t2.tl
  AND t1.aht_sec   <=> t2.aht_sec;


-- ----------------------------------------------------------------------------
-- 3. STANDARDIZE NAME CASING & WHITESPACE
--    (issue: "Name casing / whitespace", 52 rows)
--    "yuki tanaka" / "YUKI TANAKA" / "Yuki  Tanaka " -> "Yuki Tanaka"
-- ----------------------------------------------------------------------------

DELIMITER $$
DROP FUNCTION IF EXISTS title_case$$
CREATE FUNCTION title_case(input VARCHAR(100)) RETURNS VARCHAR(100)
DETERMINISTIC
BEGIN
    DECLARE result VARCHAR(100) DEFAULT '';
    DECLARE word VARCHAR(50);
    DECLARE remaining VARCHAR(100);
    DECLARE space_pos INT;
    SET remaining = TRIM(input);
    WHILE LENGTH(remaining) > 0 DO
        SET space_pos = LOCATE(' ', remaining);
        IF space_pos = 0 THEN
            SET word = remaining;
            SET remaining = '';
        ELSE
            SET word = LEFT(remaining, space_pos - 1);
            SET remaining = SUBSTRING(remaining, space_pos + 1);
        END IF;
        IF LENGTH(word) > 0 THEN
            SET result = CONCAT(result, IF(result = '', '', ' '),
                UPPER(LEFT(word,1)), LOWER(SUBSTRING(word,2)));
        END IF;
    END WHILE;
    RETURN result;
END$$
DELIMITER ;

UPDATE staging_idle_log
SET name = title_case(REGEXP_REPLACE(TRIM(name), ' +', ' '))
WHERE name IS NOT NULL;

UPDATE staging_aht_log
SET name = title_case(REGEXP_REPLACE(TRIM(name), ' +', ' '))
WHERE name IS NOT NULL;


-- ----------------------------------------------------------------------------
-- 4. STANDARDIZE REGION NAMING   (issue: "Inconsistent Region naming", 78 rows)
--    apac / Apac / " APAC" / "APAC " / "Asia Pacific" / "AP" -> APAC  (etc.)
-- ----------------------------------------------------------------------------

UPDATE staging_idle_log
SET region = CASE
    WHEN UPPER(TRIM(region)) IN ('APAC','AP','ASIA PACIFIC')     THEN 'APAC'
    WHEN UPPER(TRIM(region)) IN ('LATAM','LATIN AMERICA')        THEN 'LATAM'
    WHEN UPPER(TRIM(region)) IN ('EMEA')                         THEN 'EMEA'
    ELSE UPPER(TRIM(region))
END
WHERE region IS NOT NULL;

UPDATE staging_aht_log
SET region = CASE
    WHEN UPPER(TRIM(region)) IN ('APAC','AP','ASIA PACIFIC')     THEN 'APAC'
    WHEN UPPER(TRIM(region)) IN ('LATAM','LATIN AMERICA')        THEN 'LATAM'
    WHEN UPPER(TRIM(region)) IN ('EMEA')                         THEN 'EMEA'
    ELSE UPPER(TRIM(region))
END
WHERE region IS NOT NULL;


-- ----------------------------------------------------------------------------
-- 5. FIX TL NAME TYPOS   (issue: "TL name typos", 21 rows)
--    Known variants mapped to canonical spelling. Add rows here if you find
--    more variants during validation (step 11).
-- ----------------------------------------------------------------------------

UPDATE staging_aht_log
SET tl = CASE
    WHEN TRIM(UPPER(tl)) IN ('TL RAMIREZ','TL RAMIRES')          THEN 'TL Ramirez'
    WHEN TRIM(UPPER(tl)) IN ('TL FERREIRA','TL FEREIRA')         THEN 'TL Ferreira'
    WHEN TRIM(UPPER(tl)) IN ('TL ZIMMERMAN','TL ZIMMERMANN')     THEN 'TL Zimmerman'
    WHEN TRIM(UPPER(tl)) IN ('TL ALVAREZ','TL ALVEREZ')          THEN 'TL Alvarez'
    ELSE TRIM(tl)
END
WHERE tl IS NOT NULL;

UPDATE staging_idle_log
SET tl = CASE
    WHEN TRIM(UPPER(tl)) IN ('TL RAMIREZ','TL RAMIRES')          THEN 'TL Ramirez'
    WHEN TRIM(UPPER(tl)) IN ('TL FERREIRA','TL FEREIRA')         THEN 'TL Ferreira'
    WHEN TRIM(UPPER(tl)) IN ('TL ZIMMERMAN','TL ZIMMERMANN')     THEN 'TL Zimmerman'
    WHEN TRIM(UPPER(tl)) IN ('TL ALVAREZ','TL ALVEREZ')          THEN 'TL Alvarez'
    ELSE TRIM(tl)
END
WHERE tl IS NOT NULL;


-- ----------------------------------------------------------------------------
-- 6. STANDARDIZE COUNTRY CODE CASE   (issue: "Country code case", 13 rows)
-- ----------------------------------------------------------------------------

UPDATE staging_aht_log
SET country = UPPER(TRIM(country))
WHERE country IS NOT NULL;

UPDATE staging_idle_log
SET country = UPPER(TRIM(country))
WHERE country IS NOT NULL;


-- ----------------------------------------------------------------------------
-- 7. BACK-FILL MISSING TL / REGION FROM TEAM MEMBER LIST
--    (issues: "Missing TL" 19 rows, "Missing Country" 59 rows)
--    team_member_list has Region, not Country, so Country gaps are flagged
--    for manual review rather than guessed at.
-- ----------------------------------------------------------------------------

UPDATE staging_idle_log s
JOIN team_member_list t ON s.name = t.name
SET s.tl = t.tl
WHERE s.tl IS NULL;

UPDATE staging_aht_log s
JOIN team_member_list t ON s.name = t.name
SET s.tl = t.tl
WHERE s.tl IS NULL;

-- Country gaps that couldn't be backfilled — review manually
-- SELECT * FROM staging_aht_log WHERE country IS NULL;


-- ----------------------------------------------------------------------------
-- 8. PARSE MIXED DATE FORMATS -> proper DATE type
--    (issue: "Inconsistent date formats", 72 rows)
--    Formats seen: YYYY-MM-DD, MM/DD/YYYY, DD-Mon-YYYY, YYYY.MM.DD,
--    "Month DD, YYYY"
-- ----------------------------------------------------------------------------

ALTER TABLE staging_idle_log ADD COLUMN clean_date DATE;
ALTER TABLE staging_aht_log  ADD COLUMN clean_date DATE;

UPDATE staging_idle_log
SET clean_date = CASE
    WHEN raw_date REGEXP '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'            THEN STR_TO_DATE(raw_date, '%Y-%m-%d')
    WHEN raw_date REGEXP '^[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}$'        THEN STR_TO_DATE(raw_date, '%m/%d/%Y')
    WHEN raw_date REGEXP '^[0-9]{1,2}-[A-Za-z]{3}-[0-9]{4}$'       THEN STR_TO_DATE(raw_date, '%d-%b-%Y')
    WHEN raw_date REGEXP '^[0-9]{4}\\.[0-9]{2}\\.[0-9]{2}$'        THEN STR_TO_DATE(raw_date, '%Y.%m.%d')
    WHEN raw_date REGEXP '^[A-Za-z]+ [0-9]{1,2}, [0-9]{4}$'        THEN STR_TO_DATE(raw_date, '%M %d, %Y')
    ELSE NULL
END;

UPDATE staging_aht_log
SET clean_date = CASE
    WHEN raw_date REGEXP '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'            THEN STR_TO_DATE(raw_date, '%Y-%m-%d')
    WHEN raw_date REGEXP '^[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}$'        THEN STR_TO_DATE(raw_date, '%m/%d/%Y')
    WHEN raw_date REGEXP '^[0-9]{1,2}-[A-Za-z]{3}-[0-9]{4}$'       THEN STR_TO_DATE(raw_date, '%d-%b-%Y')
    WHEN raw_date REGEXP '^[0-9]{4}\\.[0-9]{2}\\.[0-9]{2}$'        THEN STR_TO_DATE(raw_date, '%Y.%m.%d')
    WHEN raw_date REGEXP '^[A-Za-z]+ [0-9]{1,2}, [0-9]{4}$'        THEN STR_TO_DATE(raw_date, '%M %d, %Y')
    ELSE NULL
END;


-- ----------------------------------------------------------------------------
-- 9. FIX IDLE RATE FORMAT + OUT-OF-RANGE VALUES
--    (issues: "Idle Rate as percent-string" 26 rows,
--              "Idle Rate out of range" 23 rows)
-- ----------------------------------------------------------------------------

ALTER TABLE staging_idle_log ADD COLUMN clean_idle_rate DECIMAL(6,4);

UPDATE staging_idle_log
SET clean_idle_rate = CASE
    WHEN idle_task_rate LIKE '%\%' THEN CAST(REPLACE(idle_task_rate,'%','') AS DECIMAL(8,2)) / 100
    WHEN idle_task_rate REGEXP '^[0-9.]+$' THEN CAST(idle_task_rate AS DECIMAL(6,4))
    ELSE NULL
END;

-- Idle Rate can't physically exceed 100% — flag rather than silently cap
UPDATE staging_idle_log
SET clean_idle_rate = NULL
WHERE clean_idle_rate > 1.0;


-- ----------------------------------------------------------------------------
-- 10. FIX AHT WRONG UNIT + PLACEHOLDER TEXT
--     (issues: "AHT stored in wrong unit" 54 rows,
--               "AHT placeholder text" 38 rows)
--     Values under 30 are implausible as seconds — almost certainly minutes
--     entered directly into a seconds field.
-- ----------------------------------------------------------------------------

ALTER TABLE staging_aht_log ADD COLUMN clean_aht_sec DECIMAL(10,2);

UPDATE staging_aht_log
SET clean_aht_sec = CASE
    WHEN aht_sec REGEXP '^[0-9]+(\\.[0-9]+)?$' THEN
        CASE WHEN CAST(aht_sec AS DECIMAL(10,2)) < 30
             THEN CAST(aht_sec AS DECIMAL(10,2)) * 60
             ELSE CAST(aht_sec AS DECIMAL(10,2))
        END
    ELSE NULL   -- catches 'N/A', 'TBD', '-', 'ERR: DIV/0'
END;


-- ----------------------------------------------------------------------------
-- 11. VALIDATION — run these and review before trusting the output
-- ----------------------------------------------------------------------------

-- Row counts (compare against RAW: 526 idle / 1,009 AHT before cleaning)
SELECT COUNT(*) AS idle_rows_remaining FROM staging_idle_log;
SELECT COUNT(*) AS aht_rows_remaining  FROM staging_aht_log;

-- Dates that failed to parse — should be empty; investigate any hits
SELECT * FROM staging_idle_log WHERE raw_date IS NOT NULL AND clean_date IS NULL;
SELECT * FROM staging_aht_log  WHERE raw_date IS NOT NULL AND clean_date IS NULL;

-- Idle rates that were dropped as impossible (>100%) — should be empty after review
SELECT * FROM staging_idle_log WHERE idle_task_rate IS NOT NULL AND clean_idle_rate IS NULL;

-- Region / TL values should now be a short, clean list — eyeball these
SELECT DISTINCT region FROM staging_idle_log;
SELECT DISTINCT tl     FROM staging_aht_log;

-- Any remaining Country gaps that couldn't be backfilled
SELECT COUNT(*) AS unresolved_country_gaps FROM staging_aht_log WHERE country IS NULL;


-- ----------------------------------------------------------------------------
-- 12. FINAL CLEAN OUTPUT TABLES — connect Tableau to these
-- ----------------------------------------------------------------------------

DROP TABLE IF EXISTS clean_idle_log;
CREATE TABLE clean_idle_log AS
SELECT
    clean_date          AS `date`,
    week,
    region,
    country,
    tl,
    name,
    audit_stage,
    product_level_1,
    product_level_2,
    clean_idle_rate      AS idle_task_rate,
    avg_idle_duration,
    max_idle_duration
FROM staging_idle_log
WHERE clean_date IS NOT NULL;

DROP TABLE IF EXISTS clean_aht_log;
CREATE TABLE clean_aht_log AS
SELECT
    clean_date          AS `date`,
    week,
    region,
    country,
    tl,
    name,
    audit_stage,
    product_level_1,
    product_level_2,
    clean_aht_sec         AS aht_sec,
    aht_target_sec,
    meets_target
FROM staging_aht_log
WHERE clean_date IS NOT NULL;

-- Done. Connect Tableau (New Data Source -> MySQL) to clean_idle_log
-- and clean_aht_log, relate them on Name + Date, and build from there.
