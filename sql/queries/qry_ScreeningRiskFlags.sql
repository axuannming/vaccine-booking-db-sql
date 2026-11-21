SELECT
    S.ScreeningID,
    C.CitizenName,
    A.AppointmentDate,
    S.Temperature,
    S.RiskFlag,
    S.AllergyConfirmed,
    S.ContraindicationNote
FROM (PreVaccinationScreening AS S
INNER JOIN Citizens AS C ON S.CitizenID = C.CitizenID)
INNER JOIN Appointments AS A ON S.AppointmentID = A.AppointmentID
WHERE S.RiskFlag = True
   OR S.Temperature >= 37.5
   OR S.AllergyConfirmed = True
ORDER BY S.ScreeningDate DESC, C.CitizenName;

