SELECT
tblClients.Name AS [Client Name],
COUNT(tblOrders.Article) AS [Orders],
SUM(tblOrders.Payed) AS [Payed]
FROM
tblOrders
INNER JOIN
tblClients ON tblOrders.Client = tblClients.ID
GROUP BY
tblClients.Name
ORDER BY
SUM(tblOrders.Payed) DESC
