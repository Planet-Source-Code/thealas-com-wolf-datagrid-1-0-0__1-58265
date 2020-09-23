SELECT 
tblOrders.ID AS ID_tblOrders,
tblArticles.ID AS ID_tblArticles,
tblClients.ID AS ID_tblClients,
tblEmployees.ID AS ID_tblEmployees,
tblEmployees.Name AS [Employee],
tblClients.Name AS [Client], 
tblArticles.Article,
tblOrders.Payed,
tblOrders.Owing,
tblOrders.Demands,
tblOrders.Rebate,
tblOrders.TP,
tblOrders.CP,
tblOrders.[Shipment Date],
tblOrders.[Payment Date]
FROM
((tblOrders INNER JOIN tblClients ON tblOrders.Client = tblClients.ID)
INNER JOIN tblArticles ON tblOrders.Article = tblArticles.ID)
INNER JOIN tblEmployees ON tblClients.Employee = tblEmployees.ID
WHERE tblClients.Name LIKE '%'
ORDER BY tblOrders.[ID] ASC



