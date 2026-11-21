SELECT
    V.VaccineName,
    SE.SeverityLevel,
    Count(SE.ReportID) AS ReportCount
FROM ((VaccinationRecords AS VR
INNER JOIN Vaccines AS V ON VR.VaccineID = V.VaccineID)
INNER JOIN SideEffects AS SE ON VR.RecordID = SE.VaccinationRecordID)
GROUP BY V.VaccineName, SE.SeverityLevel
ORDER BY V.VaccineName, SE.SeverityLevel;

