SELECT
    A.CitizenID,
    C.CitizenName,
    DateValue([A].[AppointmentDate]) AS ConflictDate,
    Count(A.AppointmentID) AS ConflictCount,
    Min(A.AppointmentID) AS FirstAppointmentID,
    Max(A.AppointmentID) AS LastAppointmentID
FROM Appointments AS A
INNER JOIN Citizens AS C ON A.CitizenID = C.CitizenID
WHERE A.AppointmentStatus <> 'Completed'
GROUP BY A.CitizenID, C.CitizenName, DateValue([A].[AppointmentDate])
HAVING Count(A.AppointmentID) > 1
ORDER BY Count(A.AppointmentID) DESC, DateValue([A].[AppointmentDate]) DESC;

