-- CREATE VIEW
CREATE OR REPLACE VIEW `summary_report` AS
SELECT
    -- UUID End Date YYMMDD001 002 (count from the previous date)
    l.`Emp_ID` AS `ID Group / Cell`,
    -- The group name, such as "SAI-2". The client takes the brand from the part before "-".
    l.`Emp_Name` AS `Nama Group / Cell`,
    s.`Emp_ID` AS `ID PPSKT`,
    s.`Emp_Name` AS `Nama PPSKT`,
    ph.`ProcessH_Set` As `Set Sample`,
    ph.`ProcessH_Answer1` AS `Loose / Void`,
    ph.`ProcessH_Answer2` AS `Wrinkled`,
    ph.`ProcessH_Answer3` AS `Improper Cut 1`,
    ph.`ProcessH_Answer4` AS `Improper Cut 3`,
    ph.`ProcessH_Answer5` AS `Improper Cut 4`,
    ph.`ProcessH_Answer6` AS `Improper Cut 5`,
    ph.`ProcessH_Answer7` AS `TPIS`,
    ph.`ProcessH_Answer8` AS `Weak / Open 1`,
    ph.`ProcessH_Answer9` AS `Weak / Open 3`,
    ph.`ProcessH_Answer10` AS `Weak / Open 4`,
    ph.`ProcessH_Answer11` AS `Weak / Open 5`,
    ph.`ProcessH_Answer12` AS `Weak / Open 6`,
    ph.`ProcessH_Answer13` AS `Tear / Ragged 1`,
    ph.`ProcessH_Answer14` AS `Tear / Ragged 2`,
    ph.`ProcessH_Answer15` AS `TBU`,
    ph.`ProcessH_Answer16` AS `Pleat`,
    ph.`ProcessH_Answer17` AS `Tear / Hole`,
    ph.`ProcessH_Answer18` AS `Smeared / Dirty`,
    ph.`ProcessH_Answer19` AS `Artwork Alignment`,
    ph.`ProcessH_Answer20` AS `Incompletely Field`,
    ph.`ProcessH_Answer21` AS `Foreign Matter`,
    ph.`ProcessH_Answer22` AS `Clove Spot`,
    ph.`ProcessH_Answer23` AS `Trial 1`,
    ph.`ProcessH_Answer24` AS `Reserved 2`,
    ph.`ProcessH_Answer25` AS `Reserved 3`,
    ph.`ProcessH_Answer26` AS `Reserved 4`,
    ph.`ProcessH_Answer27` AS `Reserved 5`,
    ph.`ProcessH_Answer28` AS `Reserved 6`,
    ph.`ProcessH_Answer29` AS `Reserved 7`,
    ph.`ProcessH_Answer30` AS `Reserved 8`,
    ph.`ProcessH_Answer31` AS `Reserved 9`,
    ph.`ProcessH_Answer32` AS `Reserved 10`,

    -- UB
    ROUND(AVG(CASE WHEN (CASE ph.`scale_no` WHEN 1 THEN pr.`ProcessR_RokokKe` WHEN 2 THEN pr.`ProcessR_RokokKe` + 3 END) = 1 THEN pr.`ProcessR_UB` END),2) AS `UB [1]`,
    ROUND(AVG(CASE WHEN (CASE ph.`scale_no` WHEN 1 THEN pr.`ProcessR_RokokKe` WHEN 2 THEN pr.`ProcessR_RokokKe` + 3 END) = 2 THEN pr.`ProcessR_UB` END),2) AS `UB [2]`,
    ROUND(AVG(CASE WHEN (CASE ph.`scale_no` WHEN 1 THEN pr.`ProcessR_RokokKe` WHEN 2 THEN pr.`ProcessR_RokokKe` + 3 END) = 3 THEN pr.`ProcessR_UB` END),2) AS `UB [3]`,
    ROUND(AVG(CASE WHEN (CASE ph.`scale_no` WHEN 1 THEN pr.`ProcessR_RokokKe` WHEN 2 THEN pr.`ProcessR_RokokKe` + 3 END) = 4 THEN pr.`ProcessR_UB` END),2) AS `UB [4]`,
    ROUND(AVG(CASE WHEN (CASE ph.`scale_no` WHEN 1 THEN pr.`ProcessR_RokokKe` WHEN 2 THEN pr.`ProcessR_RokokKe` + 3 END) = 5 THEN pr.`ProcessR_UB` END),2) AS `UB [5]`,
    ROUND(AVG(CASE WHEN (CASE ph.`scale_no` WHEN 1 THEN pr.`ProcessR_RokokKe` WHEN 2 THEN pr.`ProcessR_RokokKe` + 3 END) = 6 THEN pr.`ProcessR_UB` END),2) AS `UB [6]`,

    -- UH
    ROUND(AVG(CASE WHEN (CASE ph.`scale_no` WHEN 1 THEN pr.`ProcessR_RokokKe` WHEN 2 THEN pr.`ProcessR_RokokKe` + 3 END) = 1 THEN pr.`ProcessR_UH` END),2) AS `UH [1]`,
    ROUND(AVG(CASE WHEN (CASE ph.`scale_no` WHEN 1 THEN pr.`ProcessR_RokokKe` WHEN 2 THEN pr.`ProcessR_RokokKe` + 3 END) = 2 THEN pr.`ProcessR_UH` END),2) AS `UH [2]`,
    ROUND(AVG(CASE WHEN (CASE ph.`scale_no` WHEN 1 THEN pr.`ProcessR_RokokKe` WHEN 2 THEN pr.`ProcessR_RokokKe` + 3 END) = 3 THEN pr.`ProcessR_UH` END),2) AS `UH [3]`,
    ROUND(AVG(CASE WHEN (CASE ph.`scale_no` WHEN 1 THEN pr.`ProcessR_RokokKe` WHEN 2 THEN pr.`ProcessR_RokokKe` + 3 END) = 4 THEN pr.`ProcessR_UH` END),2) AS `UH [4]`,
    ROUND(AVG(CASE WHEN (CASE ph.`scale_no` WHEN 1 THEN pr.`ProcessR_RokokKe` WHEN 2 THEN pr.`ProcessR_RokokKe` + 3 END) = 5 THEN pr.`ProcessR_UH` END),2) AS `UH [5]`,
    ROUND(AVG(CASE WHEN (CASE ph.`scale_no` WHEN 1 THEN pr.`ProcessR_RokokKe` WHEN 2 THEN pr.`ProcessR_RokokKe` + 3 END) = 6 THEN pr.`ProcessR_UH` END),2) AS `UH [6]`,

    -- Average gabungan
    ROUND(AVG(pr.`ProcessR_UB`),2) AS `Avg CW [1]`,
    ROUND(AVG(pr.`ProcessR_UH`),2) AS `Avg CW [2]`,

    MIN(ph.`start_at`) AS `Start`,
    MAX(ph.`end_at`) AS `End`,
    DATE(MIN(ph.`created_at`)) AS `Tanggal`

FROM process_header ph
JOIN member l ON l.`Member_ID` = ph.`ProcessH_LeaderID`
JOIN member s ON s.`Member_ID` = ph.`ProcessH_StaffID`
JOIN process_raw pr ON pr.`ProcessR_ProcessHID` = ph.`ProcessH_ID`

GROUP BY
    l.`Emp_ID`,
    l.`Emp_Name`,
    s.`Emp_ID`,
    s.`Emp_Name`,
    ph.`start_at`;
