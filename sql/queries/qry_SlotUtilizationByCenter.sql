SELECT
    VC.CenterName,
    Count(S.SlotID) AS TotalSlots,
    Sum(S.BookedCount) AS TotalBookings,
    Sum(S.Capacity) AS TotalCapacity,
    IIf(Sum(S.Capacity)=0,0,Sum(S.BookedCount)*100.0/Sum(S.Capacity)) AS UtilizationPercent
FROM AppointmentSlots AS S
INNER JOIN VaccinationCenters AS VC ON S.CenterID = VC.CenterID
GROUP BY VC.CenterName
HAVING Count(S.SlotID) > 0
ORDER BY 5 DESC, VC.CenterName;

