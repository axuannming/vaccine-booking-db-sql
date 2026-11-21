SELECT
    C.District,
    Count(C.CitizenID) AS TotalCitizens,
    Sum(IIf(Not IsNull(VRC.CitizenID),1,0)) AS VaccinatedCitizens,
    IIf(Count(C.CitizenID)=0,0,Sum(IIf(Not IsNull(VRC.CitizenID),1,0))*100.0/Count(C.CitizenID)) AS VaccinationRatePercent
FROM Citizens AS C
LEFT JOIN (
    SELECT CitizenID
    FROM VaccinationRecords
    GROUP BY CitizenID
)  AS VRC ON C.CitizenID = VRC.CitizenID
GROUP BY C.District
ORDER BY C.District;

