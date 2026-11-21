SELECT
    A.AppointmentID,
    A.CitizenID,
    C.HKID,
    C.CitizenName,
    A.VaccineID,
    V.VaccineName,
    A.CenterID,
    VC.CenterName,
    A.SlotID,
    S.SlotDate,
    Format([S].[SlotTime],'hh:nn AM/PM') AS SlotTimeLabel,
    A.BatchID,
    B.BatchNumber,
    A.DoseNo,
    A.AppointmentDate,
    A.AppointmentStatus,
    A.ActualVaccinationDate,
    A.VaccinatedBy,
    U.Username AS VaccinatedByName,
    A.Notes,
    A.CancellationReason
FROM ((((((Appointments AS A
INNER JOIN Citizens AS C ON A.CitizenID = C.CitizenID)
INNER JOIN Vaccines AS V ON A.VaccineID = V.VaccineID)
INNER JOIN VaccinationCenters AS VC ON A.CenterID = VC.CenterID)
LEFT JOIN AppointmentSlots AS S ON A.SlotID = S.SlotID)
LEFT JOIN VaccineBatches AS B ON A.BatchID = B.BatchID)
LEFT JOIN Users AS U ON A.VaccinatedBy = U.UserID)
ORDER BY A.AppointmentDate DESC, A.AppointmentID DESC;

