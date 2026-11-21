SELECT
    VC.CenterName,
    V.VaccineName,
    Sum(IIf([T].[TransactionType]='Used',[T].[Quantity],0)) AS UsedQty,
    Sum(IIf([T].[TransactionType]='Wastage',[T].[Quantity],0)) AS WastageQty,
    Sum(IIf([T].[TransactionType]='Adjustment',[T].[Quantity],0)) AS AdjustmentQty
FROM ((InventoryTransactions AS T
INNER JOIN VaccineBatches AS B ON T.BatchID = B.BatchID)
INNER JOIN Vaccines AS V ON B.VaccineID = V.VaccineID)
INNER JOIN VaccinationCenters AS VC ON T.CenterID = VC.CenterID
GROUP BY VC.CenterName, V.VaccineName
HAVING Sum(IIf([T].[TransactionType]='Used',[T].[Quantity],0)) > 0
    OR Sum(IIf([T].[TransactionType]='Wastage',[T].[Quantity],0)) > 0
ORDER BY 4 DESC, 3 DESC, VC.CenterName;

