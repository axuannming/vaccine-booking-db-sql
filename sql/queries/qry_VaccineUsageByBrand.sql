SELECT
    V.VaccineName,
    Count(VR.RecordID) AS VaccinationCount
FROM Vaccines AS V
LEFT JOIN VaccinationRecords AS VR ON V.VaccineID = VR.VaccineID
GROUP BY V.VaccineName
HAVING Count(VR.RecordID) > 0
ORDER BY Count(VR.RecordID) DESC, V.VaccineName;

