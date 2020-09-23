SELECT
tblClients.Name AS [Client Name],
COUNT(tblOrders.Article) AS [Orders],
SUM(tblOrders.Payed) AS [Payed]
FROM
tblOrders
INNER JOIN
tblClients ON tblOrders.Client = tblClients.ID
WHERE tblOrders.[%KOLONA%] >= #%DATUM1%# AND tblOrders.[%KOLONA%] < #%DATUM2%#
GROUP BY
tblClients.Name
ORDER BY
SUM(tblOrders.Payed) DESC

