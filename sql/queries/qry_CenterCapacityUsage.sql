SELECT
    VC.CenterName,
    S.SlotDate,
    Format([S].[SlotTime],'hh:nn AM/PM') AS SlotTimeLabel,
    S.Capacity,
    S.BookedCount,
    IIf([S].[Capacity]=0,0,([S].[BookedCount]*100.0)/[S].[Capacity]) AS UtilizationPercent,
    S.SlotStatus
FROM AppointmentSlots AS S
INNER JOIN VaccinationCenters AS VC ON S.CenterID = VC.CenterID
ORDER BY S.SlotDate, VC.CenterName, S.SlotTime;

