SELECT
    X.AgeGroup,
    Count(X.RecordID) AS VaccinationRecords,
    Sum(IIf(Not IsNull(X.ReportID),1,0)) AS SideEffectReports,
    IIf(Count(X.RecordID)=0,0,Sum(IIf(Not IsNull(X.ReportID),1,0))*100.0/Count(X.RecordID)) AS SideEffectRatePercent
FROM (
    SELECT
        VR.RecordID,
        IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=11,"0-11",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=17,"12-17",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=29,"18-29",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=44,"30-44",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=59,"45-59","60+"))))) AS AgeGroup,
        SE.ReportID
    FROM (VaccinationRecords AS VR
    INNER JOIN Citizens ON VR.CitizenID = Citizens.CitizenID)
    LEFT JOIN SideEffects AS SE ON VR.RecordID = SE.VaccinationRecordID
)  AS X
GROUP BY X.AgeGroup
ORDER BY X.AgeGroup;

