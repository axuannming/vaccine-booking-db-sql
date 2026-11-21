SELECT
    VR.CitizenID,
    C.CitizenName,
    V.VaccineName,
    VR.VaccinationDate AS FirstDoseDate,
    DateAdd('d', V.DoseIntervalDays, VR.VaccinationDate) AS RecommendedSecondDoseDate,
    DateDiff('d', DateAdd('d', V.DoseIntervalDays, VR.VaccinationDate), Date()) AS DaysOverdue
FROM (VaccinationRecords AS VR
INNER JOIN Citizens AS C ON VR.CitizenID = C.CitizenID)
INNER JOIN Vaccines AS V ON VR.VaccineID = V.VaccineID
WHERE VR.DoseNo = 1
  AND V.RequiredDoses >= 2
  AND DateAdd('d', V.DoseIntervalDays, VR.VaccinationDate) <= Date()
  AND DCount('*','VaccinationRecords','CitizenID=' & [VR].[CitizenID] & ' AND VaccineID=' & [VR].[VaccineID] & ' AND DoseNo>=2') = 0
ORDER BY 6 DESC, 2;

