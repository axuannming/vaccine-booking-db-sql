SELECT
    C.CitizenID,
    C.CitizenName,
    A.AppointmentID,
    A.AppointmentDate,
    V.VaccineName,
    VR.RecordID,
    VR.DoseNo,
    VR.VaccinationDate,
    VC.CenterName,
    B.BatchNumber,
    O.ImmediateReaction,
    O.ObservationOutcome,
    SE.SymptomType,
    SE.SeverityLevel,
    SE.ReportDate,
    SE.FollowUpStatus
FROM (((((((Citizens AS C
LEFT JOIN Appointments AS A ON C.CitizenID = A.CitizenID)
LEFT JOIN VaccinationRecords AS VR ON A.AppointmentID = VR.AppointmentID)
LEFT JOIN Vaccines AS V ON VR.VaccineID = V.VaccineID)
LEFT JOIN VaccinationCenters AS VC ON A.CenterID = VC.CenterID)
LEFT JOIN VaccineBatches AS B ON VR.BatchID = B.BatchID)
LEFT JOIN ObservationRecords AS O ON VR.RecordID = O.RecordID)
LEFT JOIN SideEffects AS SE ON VR.RecordID = SE.VaccinationRecordID)
ORDER BY C.CitizenID, VR.VaccinationDate;

