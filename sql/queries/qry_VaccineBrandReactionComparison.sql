SELECT
    V.VaccineName,
    Count(VR.RecordID) AS VaccinationRecords,
    Sum(IIf(Not IsNull(SE.ReportID),1,0)) AS SideEffectReports,
    IIf(Count(VR.RecordID)=0,0,Sum(IIf(Not IsNull(SE.ReportID),1,0))*100.0/Count(VR.RecordID)) AS ReactionRatePercent
FROM ((VaccinationRecords AS VR
INNER JOIN Vaccines AS V ON VR.VaccineID = V.VaccineID)
LEFT JOIN SideEffects AS SE ON VR.RecordID = SE.VaccinationRecordID)
GROUP BY V.VaccineName
HAVING Count(VR.RecordID) > 0
ORDER BY 4 DESC, V.VaccineName;

