
SELECT 
tblArticles.Article, 
COUNT(tblArticles.Article) AS [Orders],
SUM(tblOrders.Payed) AS [Payed],
SUM(tblOrders.CP) AS [CP],
SUM(tblOrders.Owing) AS [Owing],
SUM(tblOrders.Demands) AS [Demands]
FROM
((tblOrders INNER JOIN tblClients ON tblOrders.Client = tblClients.ID)
INNER JOIN tblArticles ON tblOrders.Article = tblArticles.ID)
INNER JOIN tblEmployees ON tblClients.Employee = tblEmployees.ID
WHERE tblClients.Name LIKE '%'
GROUP BY tblArticles.Article
ORDER BY SUM(tblOrders.CP) DESC

