SELECT
    VC.CenterName,
    Count(A.AppointmentID) AS TotalAppointments,
    Sum(IIf([A].[AppointmentStatus]='NoShow',1,0)) AS NoShowCount,
    IIf(Count(A.AppointmentID)=0,0,Sum(IIf([A].[AppointmentStatus]='NoShow',1,0))*100.0/Count(A.AppointmentID)) AS NoShowRatePercent
FROM VaccinationCenters AS VC
LEFT JOIN Appointments AS A ON VC.CenterID = A.CenterID
GROUP BY VC.CenterName
HAVING Count(A.AppointmentID) > 0
ORDER BY 4 DESC, VC.CenterName;

