-- Index
CREATE INDEX `ix_ph_hid` ON `process_header` (`ProcessH_ID`);
CREATE INDEX `ix_pt_hid` ON `process_result` (`ProcessT_ProcessHID`)
CREATE INDEX `ix_pr_hid_cover` ON `process_raw` (`ProcessR_ProcessHID`, `ProcessR_RokokKe`, `ProcessR_UB`, `ProcessR_UH`);
CREATE INDEX `ix_ph_leader` ON `process_header` (`ProcessH_LeaderID`);
CREATE INDEX `ix_ph_staff` ON `process_header` (`ProcessH_StaffID`);