SELECT 
tblArticles.Article, 
COUNT(tblArticles.Article) AS [Orders],
SUM(tblOrders.Payed) AS [Payed]
FROM tblOrders
INNER JOIN tblArticles
ON tblOrders.Article=tblArticles.ID
GROUP BY tblArticles.Article
ORDER BY COUNT(tblArticles.Article) DESC
