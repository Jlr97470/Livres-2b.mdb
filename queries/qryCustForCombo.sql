SELECT
  tblCustomers.CustomerID,
  [LastName] & ", " & [FirstName] & " " & [MiddleInit] & "." AS CName,
  tblCustomers.City,
  tblCustomers.StateOrProvince
FROM
  tblCustomers
ORDER BY
  [LastName] & ", " & [FirstName] & " " & [MiddleInit] & ".";
