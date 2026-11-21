SELECT
    F.FollowUpID,
    C.CitizenName,
    V.VaccineName,
    SE.SymptomType,
    SE.SeverityLevel,
    F.FollowUpDate,
    F.ActionTaken,
    F.FollowUpOutcome,
    F.EscalationFlag
FROM ((((AdverseEffectFollowUps AS F
INNER JOIN SideEffects AS SE ON F.ReportID = SE.ReportID)
INNER JOIN Citizens AS C ON SE.CitizenID = C.CitizenID)
LEFT JOIN VaccinationRecords AS VR ON SE.VaccinationRecordID = VR.RecordID)
LEFT JOIN Vaccines AS V ON VR.VaccineID = V.VaccineID)
WHERE F.FollowUpOutcome IN ('Open','Monitoring','Needs Review')
   OR F.EscalationFlag = True
ORDER BY F.EscalationFlag DESC, F.FollowUpDate DESC;

