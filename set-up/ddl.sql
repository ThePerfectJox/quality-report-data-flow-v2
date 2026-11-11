-- Initial database schema for MariaDB 10.4+
-- Database: sampoerna
-- Primary keys and foreign keys are declared with the tables.
-- Reference/seed data is inserted after all schema objects are created.

SET NAMES utf8mb4;
SET SQL_MODE = 'NO_AUTO_VALUE_ON_ZERO';
SET time_zone = '+00:00';
SET FOREIGN_KEY_CHECKS = 0;

DROP DATABASE IF EXISTS `sampoerna`;
CREATE DATABASE `sampoerna`
  CHARACTER SET utf8mb4
  COLLATE utf8mb4_general_ci;
USE `sampoerna`;

START TRANSACTION;

CREATE TABLE `form_setting` (
  `Form_ID` INT NOT NULL AUTO_INCREMENT,
  `Form_Title` VARCHAR(100) NOT NULL,
  `Form_Display` TINYINT UNSIGNED NOT NULL DEFAULT 0
    COMMENT '0: Hide, 1: Form 1, 2: Form 2',
  CONSTRAINT `pk_form_setting` PRIMARY KEY (`Form_ID`),
  CONSTRAINT `chk_form_setting_display`
    CHECK (`Form_Display` IN (0, 1, 2))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

CREATE TABLE `member` (
  `Member_ID` INT NOT NULL AUTO_INCREMENT,
  `Emp_ID` VARCHAR(30) NOT NULL,
  `Emp_Name` VARCHAR(100) NOT NULL,
  `Emp_Level` TINYINT UNSIGNED NOT NULL DEFAULT 2
    COMMENT '1: GL, 2: Pekerja',
  `Emp_Level_Title` VARCHAR(50) DEFAULT NULL,
  `created_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT `pk_member` PRIMARY KEY (`Member_ID`),
  CONSTRAINT `uq_member_emp_id` UNIQUE (`Emp_ID`),
  CONSTRAINT `chk_member_level` CHECK (`Emp_Level` > 0)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

CREATE TABLE `setting` (
  `Setting_ID` INT NOT NULL AUTO_INCREMENT,
  `UB_Min` DECIMAL(8,3) NOT NULL,
  `UB_Max` DECIMAL(8,3) NOT NULL,
  `UH_Min` DECIMAL(8,3) NOT NULL,
  `UH_Max` DECIMAL(8,3) NOT NULL,
  `CW_Min` DECIMAL(8,3) NOT NULL,
  `CW_Max` DECIMAL(8,3) NOT NULL,
  `Deviasi_UB_Min` DECIMAL(8,3) NOT NULL,
  `Deviasi_UB_Max` DECIMAL(8,3) NOT NULL,
  `Deviasi_UH_Min` DECIMAL(8,3) NOT NULL,
  `Deviasi_UH_Max` DECIMAL(8,3) NOT NULL,
  `created_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT `pk_setting` PRIMARY KEY (`Setting_ID`),
  CONSTRAINT `chk_setting_ub_range` CHECK (`UB_Min` <= `UB_Max`),
  CONSTRAINT `chk_setting_uh_range` CHECK (`UH_Min` <= `UH_Max`),
  CONSTRAINT `chk_setting_cw_range` CHECK (`CW_Min` <= `CW_Max`),
  CONSTRAINT `chk_setting_dev_ub_range`
    CHECK (`Deviasi_UB_Min` <= `Deviasi_UB_Max`),
  CONSTRAINT `chk_setting_dev_uh_range`
    CHECK (`Deviasi_UH_Min` <= `Deviasi_UH_Max`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

CREATE TABLE `process_header` (
  `ProcessH_ID` INT NOT NULL AUTO_INCREMENT,
  `ProcessH_LeaderID` INT NOT NULL,
  `ProcessH_StaffID` INT NOT NULL,
  `ProcessH_WeightTotal` DECIMAL(10,3) NOT NULL,
  `ProcessH_Set` TINYINT UNSIGNED NOT NULL DEFAULT 1,
  `created_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `is_print` BOOLEAN NOT NULL DEFAULT FALSE,
  `is_print_summary` BOOLEAN NOT NULL DEFAULT TRUE,
  `start_at` DATETIME DEFAULT NULL,
  `end_at` DATETIME DEFAULT NULL,
  `scale_no` TINYINT UNSIGNED NOT NULL,
  `ProcessH_Answer1` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer2` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer3` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer4` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer5` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer6` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer7` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer8` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer9` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer10` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer11` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer12` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer13` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer14` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer15` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer16` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer17` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer18` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer19` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer20` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer21` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer22` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer23` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer24` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer25` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer26` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer27` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer28` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer29` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer30` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer31` TINYINT UNSIGNED DEFAULT NULL,
  `ProcessH_Answer32` TINYINT UNSIGNED DEFAULT NULL,
  CONSTRAINT `pk_process_header` PRIMARY KEY (`ProcessH_ID`),
  KEY `idx_process_header_leader` (`ProcessH_LeaderID`),
  KEY `idx_process_header_staff` (`ProcessH_StaffID`),
  KEY `idx_process_header_created_at` (`created_at`),
  CONSTRAINT `fk_process_header_leader`
    FOREIGN KEY (`ProcessH_LeaderID`) REFERENCES `member` (`Member_ID`)
    ON UPDATE CASCADE ON DELETE RESTRICT,
  CONSTRAINT `fk_process_header_staff`
    FOREIGN KEY (`ProcessH_StaffID`) REFERENCES `member` (`Member_ID`)
    ON UPDATE CASCADE ON DELETE RESTRICT,
  CONSTRAINT `chk_process_header_weight`
    CHECK (`ProcessH_WeightTotal` >= 0),
  CONSTRAINT `chk_process_header_scale`
    CHECK (`scale_no` BETWEEN 1 AND 4),
  CONSTRAINT `chk_process_header_time`
    CHECK (`end_at` IS NULL OR `start_at` IS NULL OR `end_at` >= `start_at`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

CREATE TABLE `process_raw` (
  `ProcessR_ID` INT NOT NULL AUTO_INCREMENT,
  `ProcessR_ProcessHID` INT NOT NULL,
  `ProcessR_RokokKe` TINYINT UNSIGNED NOT NULL COMMENT '1-3',
  `ProcessR_Index` TINYINT UNSIGNED NOT NULL COMMENT '1-10',
  `ProcessR_UB` DECIMAL(10,3) NOT NULL,
  `ProcessR_UH` DECIMAL(10,3) NOT NULL,
  CONSTRAINT `pk_process_raw` PRIMARY KEY (`ProcessR_ID`),
  CONSTRAINT `uq_process_raw_measurement`
    UNIQUE (`ProcessR_ProcessHID`, `ProcessR_RokokKe`, `ProcessR_Index`),
  KEY `idx_process_raw_header` (`ProcessR_ProcessHID`),
  CONSTRAINT `fk_process_raw_header`
    FOREIGN KEY (`ProcessR_ProcessHID`) REFERENCES `process_header` (`ProcessH_ID`)
    ON UPDATE CASCADE ON DELETE CASCADE,
  CONSTRAINT `chk_process_raw_rokkok`
    CHECK (`ProcessR_RokokKe` BETWEEN 1 AND 3),
  CONSTRAINT `chk_process_raw_index`
    CHECK (`ProcessR_Index` BETWEEN 1 AND 10)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

CREATE TABLE `process_result` (
  `ProcessT_ID` INT NOT NULL AUTO_INCREMENT,
  `ProcessT_ProcessHID` INT NOT NULL,
  `ProcessT_RokokKe` TINYINT UNSIGNED NOT NULL COMMENT '1-3',
  `ProcessT_UBAvg` DECIMAL(10,3) NOT NULL,
  `ProcessT_UHAvg` DECIMAL(10,3) NOT NULL,
  `created_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT `pk_process_result` PRIMARY KEY (`ProcessT_ID`),
  CONSTRAINT `uq_process_result_rokkok`
    UNIQUE (`ProcessT_ProcessHID`, `ProcessT_RokokKe`),
  KEY `idx_process_result_header` (`ProcessT_ProcessHID`),
  CONSTRAINT `fk_process_result_header`
    FOREIGN KEY (`ProcessT_ProcessHID`) REFERENCES `process_header` (`ProcessH_ID`)
    ON UPDATE CASCADE ON DELETE CASCADE,
  CONSTRAINT `chk_process_result_rokkok`
    CHECK (`ProcessT_RokokKe` BETWEEN 1 AND 3)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- The report is a view, not a physical table.
CREATE OR REPLACE VIEW `summary_report` AS
SELECT
  leader.`Emp_ID` AS `ID GL`,
  leader.`Emp_Name` AS `Nama GL`,
  staff.`Emp_ID` AS `ID PPSKT`,
  staff.`Emp_Name` AS `Nama PPSKT`,
  ROUND(MAX(CASE WHEN ph.`scale_no` = 1 THEN ph.`ProcessH_WeightTotal` END), 2) AS `UB [1]`,
  ROUND(MAX(CASE WHEN ph.`scale_no` = 2 THEN ph.`ProcessH_WeightTotal` END), 2) AS `UB [2]`,
  ROUND(MAX(CASE WHEN ph.`scale_no` = 3 THEN ph.`ProcessH_WeightTotal` END), 2) AS `UB [3]`,
  ROUND(MAX(CASE WHEN ph.`scale_no` = 4 THEN ph.`ProcessH_WeightTotal` END), 2) AS `UB [4]`,
  NULL AS `UB [5]`,
  NULL AS `UB [6]`,
  ROUND(MAX(CASE WHEN pr.`ProcessT_RokokKe` = 1 THEN pr.`ProcessT_UHAvg` END), 2) AS `UH [1]`,
  ROUND(MAX(CASE WHEN pr.`ProcessT_RokokKe` = 2 THEN pr.`ProcessT_UHAvg` END), 2) AS `UH [2]`,
  ROUND(MAX(CASE WHEN pr.`ProcessT_RokokKe` = 3 THEN pr.`ProcessT_UHAvg` END), 2) AS `UH [3]`,
  NULL AS `UH [4]`,
  NULL AS `UH [5]`,
  NULL AS `UH [6]`,
  ROUND(AVG(CASE WHEN ph.`ProcessH_Set` = 1 THEN ph.`ProcessH_WeightTotal` END), 2) AS `Avg CW [1]`,
  ROUND(AVG(CASE WHEN ph.`ProcessH_Set` = 2 THEN ph.`ProcessH_WeightTotal` END), 2) AS `Avg CW [2]`,
  MIN(ph.`start_at`) AS `Start`,
  MAX(ph.`end_at`) AS `End`,
  DATE(ph.`created_at`) AS `Tanggal`
FROM `process_header` ph
JOIN `member` leader ON leader.`Member_ID` = ph.`ProcessH_LeaderID`
JOIN `member` staff ON staff.`Member_ID` = ph.`ProcessH_StaffID`
LEFT JOIN `process_result` pr ON pr.`ProcessT_ProcessHID` = ph.`ProcessH_ID`
GROUP BY
  leader.`Emp_ID`, leader.`Emp_Name`,
  staff.`Emp_ID`, staff.`Emp_Name`,
  DATE(ph.`created_at`);
  
COMMIT;
SET FOREIGN_KEY_CHECKS = 1;
