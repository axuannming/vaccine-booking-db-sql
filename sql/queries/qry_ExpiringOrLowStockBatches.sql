SELECT
    B.BatchID,
    B.BatchNumber,
    V.VaccineName,
    B.ExpiryDate,
    B.CurrentStock,
    B.Supplier,
    IIf([CurrentStock] < 20, 'Low Stock', 'OK') AS StockFlag,
    IIf(DateDiff('d', Date(), [ExpiryDate]) <= 30, 'Expiring Soon', 'OK') AS ExpiryFlag
FROM VaccineBatches AS B
INNER JOIN Vaccines AS V ON B.VaccineID = V.VaccineID
WHERE B.CurrentStock < 20
   OR DateDiff('d', Date(), [ExpiryDate]) <= 30
ORDER BY B.ExpiryDate, B.CurrentStock;

