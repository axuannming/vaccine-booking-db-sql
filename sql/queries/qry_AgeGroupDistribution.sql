SELECT
    X.AgeGroup,
    Count(X.CitizenID) AS CitizenCount
FROM (
    SELECT
        Citizens.CitizenID,
        IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=11,"0-11",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=17,"12-17",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=29,"18-29",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=44,"30-44",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=59,"45-59","60+"))))) AS AgeGroup
    FROM Citizens
)  AS X
GROUP BY X.AgeGroup
ORDER BY X.AgeGroup;

