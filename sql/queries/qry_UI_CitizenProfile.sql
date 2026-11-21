SELECT
    Citizens.CitizenID,
    Citizens.HKID,
    Citizens.CitizenName,
    Citizens.FirstName,
    Citizens.LastName,
    Citizens.Gender,
    Citizens.DOB,
    DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0) AS CitizenAge,
    IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=11,"0-11",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=17,"12-17",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=29,"18-29",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=44,"30-44",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=59,"45-59","60+"))))) AS AgeGroup,
    Citizens.Phone,
    Citizens.Email,
    Citizens.[Address],
    Citizens.District,
    Citizens.Occupation,
    Citizens.PlaceOfBirth,
    Citizens.AllergyHistory
FROM Citizens
ORDER BY Citizens.CitizenID;

