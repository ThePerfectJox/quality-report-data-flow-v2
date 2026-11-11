-- =========================================================
-- Seed data. All data insertion is intentionally placed last.
-- =========================================================

INSERT INTO `form_setting` (`Form_ID`, `Form_Title`, `Form_Display`) VALUES
(1,'Loose/Void',0),(2,'Wrinkled',0),(3,'Improper Cut 01',0),(4,'Improper Cut 03',0),
(5,'Improper Cut 04',1),(6,'Improper Cut 05',1),(7,'TPIS',1),(8,'Weak/Open 1',1),
(9,'Weak/Open 3',1),(10,'Weak/Open 4',1),(11,'Weak/Open 5',1),(12,'Weak/Open 6',1),
(13,'Tear/Ragged 01',2),(14,'Tear/Ragged 02',2),(15,'TBU',2),(16,'Pleat',2),
(17,'Tear/Hole',2),(18,'Smeared/Dirty',2),(19,'Artwork Alignment',2),
(20,'Incompletely Filled',2),(21,'Foreign Matter',2),(22,'Clove Spot',2),(23,'Trial 1',0),
(24,'Reserved 2',0),(25,'Reserved 3',0),(26,'Reserved 4',0),(27,'Reserved 5',0),
(28,'Reserved 6',0),(29,'Reserved 7',0),(30,'Reserved 8',0),(31,'Reserved 9',0),
(32,'Reserved 10',0);

INSERT INTO `setting`
(`Setting_ID`,`UB_Min`,`UB_Max`,`UH_Min`,`UH_Max`,`CW_Min`,`CW_Max`,
 `Deviasi_UB_Min`,`Deviasi_UB_Max`,`Deviasi_UH_Min`,`Deviasi_UH_Max`,`created_at`)
VALUES
(1,9.8,10.2,7.8,8.2,1.88,2.16,1,1.5,1,1.5,'2026-01-23 13:25:56');