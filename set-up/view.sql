-- Result Per Header
CREATE OR REPLACE VIEW `result_per_header` AS
SELECT
    ph.`ProcessH_ID`,
    ph.`ProcessH_LeaderID`,
    ph.`ProcessH_StaffID`,
    ph.`ProcessH_Set`,

    ph.`ProcessH_Answer1`,
    ph.`ProcessH_Answer2`,
    ph.`ProcessH_Answer3`,
    ph.`ProcessH_Answer4`,
    ph.`ProcessH_Answer5`,
    ph.`ProcessH_Answer6`,
    ph.`ProcessH_Answer7`,
    ph.`ProcessH_Answer8`,
    ph.`ProcessH_Answer9`,
    ph.`ProcessH_Answer10`,
    ph.`ProcessH_Answer11`,
    ph.`ProcessH_Answer12`,
    ph.`ProcessH_Answer13`,
    ph.`ProcessH_Answer14`,
    ph.`ProcessH_Answer15`,
    ph.`ProcessH_Answer16`,
    ph.`ProcessH_Answer17`,
    ph.`ProcessH_Answer18`,
    ph.`ProcessH_Answer19`,
    ph.`ProcessH_Answer20`,
    ph.`ProcessH_Answer21`,
    ph.`ProcessH_Answer22`,
    ph.`ProcessH_Answer23`,
    ph.`ProcessH_Answer24`,
    ph.`ProcessH_Answer25`,
    ph.`ProcessH_Answer26`,
    ph.`ProcessH_Answer27`,
    ph.`ProcessH_Answer28`,
    ph.`ProcessH_Answer29`,
    ph.`ProcessH_Answer30`,
    ph.`ProcessH_Answer31`,
    ph.`ProcessH_Answer32`,

    ph.`scale_no`,
    ph.`start_at`,
    ph.`end_at`,
    ph.`created_at`,

    NULLIF(ph.`ProcessH_WeightTotal`, 0) AS `cw`,

    MAX(
        CASE
            WHEN pt.`ProcessT_RokokKe` = 1
            THEN NULLIF(pt.`ProcessT_UBAvg`, 0)
        END
    ) AS `ub_1`,

    MAX(
        CASE
            WHEN pt.`ProcessT_RokokKe` = 2
            THEN NULLIF(pt.`ProcessT_UBAvg`, 0)
        END
    ) AS `ub_2`,

    MAX(
        CASE
            WHEN pt.`ProcessT_RokokKe` = 3
            THEN NULLIF(pt.`ProcessT_UBAvg`, 0)
        END
    ) AS `ub_3`,

    MAX(
        CASE
            WHEN pt.`ProcessT_RokokKe` = 1
            THEN NULLIF(pt.`ProcessT_UHAvg`, 0)
        END
    ) AS `uh_1`,

    MAX(
        CASE
            WHEN pt.`ProcessT_RokokKe` = 2
            THEN NULLIF(pt.`ProcessT_UHAvg`, 0)
        END
    ) AS `uh_2`,

    MAX(
        CASE
            WHEN pt.`ProcessT_RokokKe` = 3
            THEN NULLIF(pt.`ProcessT_UHAvg`, 0)
        END
    ) AS `uh_3`

FROM `process_header` AS ph

LEFT JOIN `process_result` AS pt
    ON pt.`ProcessT_ProcessHID` = ph.`ProcessH_ID`

GROUP BY
    ph.`ProcessH_ID`,
    ph.`ProcessH_LeaderID`,
    ph.`ProcessH_StaffID`,
    ph.`ProcessH_Set`,

    ph.`ProcessH_Answer1`,
    ph.`ProcessH_Answer2`,
    ph.`ProcessH_Answer3`,
    ph.`ProcessH_Answer4`,
    ph.`ProcessH_Answer5`,
    ph.`ProcessH_Answer6`,
    ph.`ProcessH_Answer7`,
    ph.`ProcessH_Answer8`,
    ph.`ProcessH_Answer9`,
    ph.`ProcessH_Answer10`,
    ph.`ProcessH_Answer11`,
    ph.`ProcessH_Answer12`,
    ph.`ProcessH_Answer13`,
    ph.`ProcessH_Answer14`,
    ph.`ProcessH_Answer15`,
    ph.`ProcessH_Answer16`,
    ph.`ProcessH_Answer17`,
    ph.`ProcessH_Answer18`,
    ph.`ProcessH_Answer19`,
    ph.`ProcessH_Answer20`,
    ph.`ProcessH_Answer21`,
    ph.`ProcessH_Answer22`,
    ph.`ProcessH_Answer23`,
    ph.`ProcessH_Answer24`,
    ph.`ProcessH_Answer25`,
    ph.`ProcessH_Answer26`,
    ph.`ProcessH_Answer27`,
    ph.`ProcessH_Answer28`,
    ph.`ProcessH_Answer29`,
    ph.`ProcessH_Answer30`,
    ph.`ProcessH_Answer31`,
    ph.`ProcessH_Answer32`,

    ph.`scale_no`,
    ph.`start_at`,
    ph.`end_at`,
    ph.`created_at`,
    ph.`ProcessH_WeightTotal`;

-- Measurement Per Header
CREATE OR REPLACE VIEW `measurement_per_header` AS
SELECT
    ph.`ProcessH_ID`,
    ph.`ProcessH_LeaderID`,
    ph.`ProcessH_StaffID`,
    ph.`ProcessH_Set`,
    ph.`scale_no`,
    ph.`start_at`,
    ph.`end_at`,
    ph.`created_at`,
    NULLIF(ph.`ProcessH_WeightTotal`, 0) AS `cw`,

    -- UB measurements
    MAX(
        CASE
            WHEN pt.`ProcessT_RokokKe` = 1
            THEN NULLIF(pt.`ProcessT_UBAvg`, 0)
        END
    ) AS `ub_1`,

    MAX(
        CASE
            WHEN pt.`ProcessT_RokokKe` = 2
            THEN NULLIF(pt.`ProcessT_UBAvg`, 0)
        END
    ) AS `ub_2`,

    MAX(
        CASE
            WHEN pt.`ProcessT_RokokKe` = 3
            THEN NULLIF(pt.`ProcessT_UBAvg`, 0)
        END
    ) AS `ub_3`,

    -- UH measurements
    MAX(
        CASE
            WHEN pt.`ProcessT_RokokKe` = 1
            THEN NULLIF(pt.`ProcessT_UHAvg`, 0)
        END
    ) AS `uh_1`,

    MAX(
        CASE
            WHEN pt.`ProcessT_RokokKe` = 2
            THEN NULLIF(pt.`ProcessT_UHAvg`, 0)
        END
    ) AS `uh_2`,

    MAX(
        CASE
            WHEN pt.`ProcessT_RokokKe` = 3
            THEN NULLIF(pt.`ProcessT_UHAvg`, 0)
        END
    ) AS `uh_3`

FROM `process_header` AS ph

LEFT JOIN `process_result` AS pt
    ON pt.`ProcessT_ProcessHID` = ph.`ProcessH_ID`

GROUP BY
    ph.`ProcessH_ID`,
    ph.`ProcessH_LeaderID`,
    ph.`ProcessH_StaffID`,
    ph.`ProcessH_Set`,
    ph.`scale_no`,
    ph.`start_at`,
    ph.`end_at`,
    ph.`created_at`,
    ph.`ProcessH_WeightTotal`;

-- Summary Report
CREATE OR REPLACE VIEW `summary_report` AS
SELECT
    l.`Emp_ID` AS `ID Group / Cell`,
    l.`Emp_Name` AS `Nama Group / Cell`,
    s.`Emp_ID` AS `ID PPSKT`,
    s.`Emp_Name` AS `Nama PPSKT`,
    MAX(rph.`ProcessH_Set`) AS `Set Sample`,
    COALESCE(MAX(rph.`ProcessH_Answer1`), 0) AS `Loose / Void`,
    COALESCE(MAX(rph.`ProcessH_Answer2`), 0) AS `Wrinkled`,
    COALESCE(MAX(rph.`ProcessH_Answer3`), 0) AS `Improper Cut 1`,
    COALESCE(MAX(rph.`ProcessH_Answer4`), 0) AS `Improper Cut 3`,
    COALESCE(MAX(rph.`ProcessH_Answer5`), 0) AS `Improper Cut 4`,
    COALESCE(MAX(rph.`ProcessH_Answer6`), 0) AS `Improper Cut 5`,
    COALESCE(MAX(rph.`ProcessH_Answer7`), 0) AS `TPIS`,
    COALESCE(MAX(rph.`ProcessH_Answer8`), 0) AS `Weak / Open 1`,
    COALESCE(MAX(rph.`ProcessH_Answer9`), 0) AS `Weak / Open 3`,
    COALESCE(MAX(rph.`ProcessH_Answer10`), 0) AS `Weak / Open 4`,
    COALESCE(MAX(rph.`ProcessH_Answer11`), 0) AS `Weak / Open 5`,
    COALESCE(MAX(rph.`ProcessH_Answer12`), 0) AS `Weak / Open 6`,
    COALESCE(MAX(rph.`ProcessH_Answer13`), 0) AS `Tear / Ragged 1`,
    COALESCE(MAX(rph.`ProcessH_Answer14`), 0) AS `Tear / Ragged 2`,
    COALESCE(MAX(rph.`ProcessH_Answer15`), 0) AS `TBU`,
    COALESCE(MAX(rph.`ProcessH_Answer16`), 0) AS `Pleat`,
    COALESCE(MAX(rph.`ProcessH_Answer17`), 0) AS `Tear / Hole`,
    COALESCE(MAX(rph.`ProcessH_Answer18`), 0) AS `Smeared / Dirty`,
    COALESCE(MAX(rph.`ProcessH_Answer19`), 0) AS `Artwork Alignment`,
    COALESCE(MAX(rph.`ProcessH_Answer20`), 0) AS `Incompletely Field`,
    COALESCE(MAX(rph.`ProcessH_Answer21`), 0) AS `Foreign Matter`,
    COALESCE(MAX(rph.`ProcessH_Answer22`), 0) AS `Clove Spot`,
    COALESCE(MAX(rph.`ProcessH_Answer23`), 0) AS `Trial 1`,
    COALESCE(MAX(rph.`ProcessH_Answer24`), 0) AS `Reserved 2`,
    COALESCE(MAX(rph.`ProcessH_Answer25`), 0) AS `Reserved 3`,
    COALESCE(MAX(rph.`ProcessH_Answer26`), 0) AS `Reserved 4`,
    COALESCE(MAX(rph.`ProcessH_Answer27`), 0) AS `Reserved 5`,
    COALESCE(MAX(rph.`ProcessH_Answer28`), 0) AS `Reserved 6`,
    COALESCE(MAX(rph.`ProcessH_Answer29`), 0) AS `Reserved 7`,
    COALESCE(MAX(rph.`ProcessH_Answer30`), 0) AS `Reserved 8`,
    COALESCE(MAX(rph.`ProcessH_Answer31`), 0) AS `Reserved 9`,
    COALESCE(MAX(rph.`ProcessH_Answer32`), 0) AS `Reserved 10`,

    -- UB scales 1 and 2
    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (1, 2)
        THEN rph.`ub_1`
    END), 2) AS `UB [1]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (1, 2)
        THEN rph.`ub_2`
    END), 2) AS `UB [2]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (1, 2)
        THEN rph.`ub_3`
    END), 2) AS `UB [3]`,

    -- UB scales 3 and 4
    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (3, 4)
        THEN rph.`ub_1`
    END), 2) AS `UB [4]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (3, 4)
        THEN rph.`ub_2`
    END), 2) AS `UB [5]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (3, 4)
        THEN rph.`ub_3`
    END), 2) AS `UB [6]`,

    -- UH scales 1 and 2
    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (1, 2)
        THEN rph.`uh_1`
    END), 2) AS `UH [1]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (1, 2)
        THEN rph.`uh_2`
    END), 2) AS `UH [2]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (1, 2)
        THEN rph.`uh_3`
    END), 2) AS `UH [3]`,

    -- UH scales 3 and 4
    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (3, 4)
        THEN rph.`uh_1`
    END), 2) AS `UH [4]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (3, 4)
        THEN rph.`uh_2`
    END), 2) AS `UH [5]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (3, 4)
        THEN rph.`uh_3`
    END), 2) AS `UH [6]`,

    -- CW scales 1 and 2
    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (1, 2)
        THEN rph.`cw`
    END), 2) AS `Avg CW [1]`,

    -- CW scales 3 and 4
    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (3, 4)
        THEN rph.`cw`
    END), 2) AS `Avg CW [2]`,

    MIN(rph.`start_at`) AS `Start`,
    MAX(rph.`end_at`) AS `End`,
    DATE(MIN(rph.`created_at`)) AS `Tanggal`

FROM `result_per_header` AS rph

INNER JOIN `member` AS l
    ON l.`Member_ID` = rph.`ProcessH_LeaderID`

INNER JOIN `member` AS s
    ON s.`Member_ID` = rph.`ProcessH_StaffID`

GROUP BY
    l.`Emp_ID`,
    l.`Emp_Name`,
    s.`Emp_ID`,
    s.`Emp_Name`,
    rph.`start_at`;

-- Measurement Report
CREATE OR REPLACE VIEW `measurement_report` AS
SELECT
    l.`Emp_ID` AS `ID Group / Cell`,
    l.`Emp_Name` AS `Nama Group / Cell`,
    s.`Emp_ID` AS `ID PPSKT`,
    s.`Emp_Name` AS `Nama PPSKT`,

    MAX(rph.`ProcessH_Set`) AS `Set Sample`,

    -- UB from scale 1 and 2
    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (1, 2)
        THEN rph.`ub_1`
    END), 2) AS `UB [1]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (1, 2)
        THEN rph.`ub_2`
    END), 2) AS `UB [2]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (1, 2)
        THEN rph.`ub_3`
    END), 2) AS `UB [3]`,

    -- UB from scale 3 and 4
    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (3, 4)
        THEN rph.`ub_1`
    END), 2) AS `UB [4]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (3, 4)
        THEN rph.`ub_2`
    END), 2) AS `UB [5]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (3, 4)
        THEN rph.`ub_3`
    END), 2) AS `UB [6]`,

    -- UH from scale 1 and 2
    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (1, 2)
        THEN rph.`uh_1`
    END), 2) AS `UH [1]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (1, 2)
        THEN rph.`uh_2`
    END), 2) AS `UH [2]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (1, 2)
        THEN rph.`uh_3`
    END), 2) AS `UH [3]`,

    -- UH from scale 3 and 4
    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (3, 4)
        THEN rph.`uh_1`
    END), 2) AS `UH [4]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (3, 4)
        THEN rph.`uh_2`
    END), 2) AS `UH [5]`,

    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (3, 4)
        THEN rph.`uh_3`
    END), 2) AS `UH [6]`,

    -- CW from scale 1 and 2
    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (1, 2)
        THEN rph.`cw`
    END), 2) AS `Avg CW [1]`,

    -- CW from scale 3 and 4
    ROUND(AVG(CASE
        WHEN rph.`scale_no` IN (3, 4)
        THEN rph.`cw`
    END), 2) AS `Avg CW [2]`,

    MIN(rph.`start_at`) AS `Start`,
    MAX(rph.`end_at`) AS `End`,
    DATE(MIN(rph.`created_at`)) AS `Tanggal`

FROM `result_per_header` AS rph

INNER JOIN `member` AS l
    ON l.`Member_ID` = rph.`ProcessH_LeaderID`

INNER JOIN `member` AS s
    ON s.`Member_ID` = rph.`ProcessH_StaffID`

GROUP BY
    l.`Emp_ID`,
    l.`Emp_Name`,
    s.`Emp_ID`,
    s.`Emp_Name`,
    rph.`start_at`;