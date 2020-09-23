SELECT
tblClients.ID AS ID_tblClients,
tblEmployees.ID AS ID_tblEmployees,
tblEmployees.Name AS Employee,
tblClients.Name,
tblClients.Address,
tblClients.Rebate,
tblClients.Phone1,
tblClients.Phone2,
tblClients.Phone3
FROM
(tblClients INNER JOIN tblEmployees
ON tblClients.Employee = tblEmployees.ID)

