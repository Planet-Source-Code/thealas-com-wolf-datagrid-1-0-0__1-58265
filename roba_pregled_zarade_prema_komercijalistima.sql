
SELECT
tblEmployees.Name AS [Employee],
COUNT(tblOrders.Article) AS [Orders],
SUM(tblOrders.Payed) AS [Payed],
SUM(tblOrders.Owing) AS [Owing],
SUM(tblOrders.Demands) AS [Demands],
SUM(tblOrders.CP) AS [Sold CP]
FROM
(tblOrders 
INNER JOIN tblClients ON tblOrders.Client = tblClients.ID)
INNER JOIN tblEmployees ON tblClients.Employee = tblEmployees.ID
WHERE tblOrders.[%KOLONA%] >= #%DATUM1%# AND tblOrders.[%KOLONA%] < #%DATUM2%#
GROUP BY
tblEmployees.Name
ORDER BY
SUM(tblOrders.Payed) DESC
