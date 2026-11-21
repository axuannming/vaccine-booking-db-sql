SELECT
    O.ObservationID,
    C.CitizenName,
    V.VaccineName,
    VR.DoseNo,
    O.ObservationStart,
    O.ObservationEnd,
    O.ObservationOutcome
FROM (((ObservationRecords AS O
INNER JOIN VaccinationRecords AS VR ON O.RecordID = VR.RecordID)
INNER JOIN Citizens AS C ON O.CitizenID = C.CitizenID)
INNER JOIN Vaccines AS V ON VR.VaccineID = V.VaccineID)
WHERE O.ImmediateReaction = True
ORDER BY O.ObservationStart DESC, C.CitizenName;

