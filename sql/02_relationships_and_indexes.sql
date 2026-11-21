CREATE UNIQUE INDEX idx_Citizens_HKID ON Citizens (HKID);
CREATE UNIQUE INDEX idx_Users_Username ON Users (Username);
CREATE UNIQUE INDEX idx_Batches_BatchNumber ON VaccineBatches (BatchNumber);
CREATE INDEX idx_Appointments_CitizenID ON Appointments (CitizenID);
CREATE INDEX idx_Appointments_SlotID ON Appointments (SlotID);
CREATE INDEX idx_VaccinationRecords_AppointmentID ON VaccinationRecords (AppointmentID);
CREATE INDEX idx_VaccinationRecords_CitizenID ON VaccinationRecords (CitizenID);
CREATE INDEX idx_Screening_AppointmentID ON PreVaccinationScreening (AppointmentID);
CREATE INDEX idx_Consent_AppointmentID ON ConsentRecords (AppointmentID);
CREATE INDEX idx_Observation_RecordID ON ObservationRecords (RecordID);
CREATE INDEX idx_SideEffects_RecordID ON SideEffects (VaccinationRecordID);
CREATE INDEX idx_FollowUps_ReportID ON AdverseEffectFollowUps (ReportID);
CREATE INDEX idx_Notifications_CitizenID ON Notifications (CitizenID);
CREATE INDEX idx_Inventory_BatchID ON InventoryTransactions (BatchID);

