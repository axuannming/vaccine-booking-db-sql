SELECT TOP 5
    VC.CenterName,
    Count(VR.RecordID) AS TotalVaccinations
FROM (VaccinationCenters AS VC
INNER JOIN Appointments AS A ON VC.CenterID = A.CenterID)
INNER JOIN VaccinationRecords AS VR ON A.AppointmentID = VR.AppointmentID
GROUP BY VC.CenterName
ORDER BY Count(VR.RecordID) DESC, VC.CenterName;

