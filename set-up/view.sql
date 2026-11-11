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
    COALESCE(ph.`ProcessH_Answer1`, 0) AS `Loose / Void`,
    COALESCE(ph.`ProcessH_Answer2`, 0) AS `Wrinkled`,
    COALESCE(ph.`ProcessH_Answer3`, 0) AS `Improper Cut 1`,
    COALESCE(ph.`ProcessH_Answer4`, 0) AS `Improper Cut 3`,
    COALESCE(ph.`ProcessH_Answer5`, 0) AS `Improper Cut 4`,
    COALESCE(ph.`ProcessH_Answer6`, 0) AS `Improper Cut 5`,
    COALESCE(ph.`ProcessH_Answer7`, 0) AS `TPIS`,
    COALESCE(ph.`ProcessH_Answer8`, 0) AS `Weak / Open 1`,
    COALESCE(ph.`ProcessH_Answer9`, 0) AS `Weak / Open 3`,
    COALESCE(ph.`ProcessH_Answer10`, 0) AS `Weak / Open 4`,
    COALESCE(ph.`ProcessH_Answer11`, 0) AS `Weak / Open 5`,
    COALESCE(ph.`ProcessH_Answer12`, 0) AS `Weak / Open 6`,
    COALESCE(ph.`ProcessH_Answer13`, 0) AS `Tear / Ragged 1`,
    COALESCE(ph.`ProcessH_Answer14`, 0) AS `Tear / Ragged 2`,
    COALESCE(ph.`ProcessH_Answer15`, 0) AS `TBU`,
    COALESCE(ph.`ProcessH_Answer16`, 0) AS `Pleat`,
    COALESCE(ph.`ProcessH_Answer17`, 0) AS `Tear / Hole`,
    COALESCE(ph.`ProcessH_Answer18`, 0) AS `Smeared / Dirty`,
    COALESCE(ph.`ProcessH_Answer19`, 0) AS `Artwork Alignment`,
    COALESCE(ph.`ProcessH_Answer20`, 0) AS `Incompletely Field`,
    COALESCE(ph.`ProcessH_Answer21`, 0) AS `Foreign Matter`,
    COALESCE(ph.`ProcessH_Answer22`, 0) AS `Clove Spot`,
    COALESCE(ph.`ProcessH_Answer23`, 0) AS `Trial 1`,
    COALESCE(ph.`ProcessH_Answer24`, 0) AS `Reserved 2`,
    COALESCE(ph.`ProcessH_Answer25`, 0) AS `Reserved 3`,
    COALESCE(ph.`ProcessH_Answer26`, 0) AS `Reserved 4`,
    COALESCE(ph.`ProcessH_Answer27`, 0) AS `Reserved 5`,
    COALESCE(ph.`ProcessH_Answer28`, 0) AS `Reserved 6`,
    COALESCE(ph.`ProcessH_Answer29`, 0) AS `Reserved 7`,
    COALESCE(ph.`ProcessH_Answer30`, 0) AS `Reserved 8`,
    COALESCE(ph.`ProcessH_Answer31`, 0) AS `Reserved 9`,
    COALESCE(ph.`ProcessH_Answer32`, 0) AS `Reserved 10`,

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
