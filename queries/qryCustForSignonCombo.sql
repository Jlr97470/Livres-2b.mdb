SELECT
  [LastName] & ", " & [FirstName] & " " & [MiddleInit] & "." AS CustName,
  tblCustomers.CustomerID,
  tblCustomers.Password,
  tblCustomers.Admin
FROM
  tblCustomers
ORDER BY
  [LastName] & ", " & [FirstName] & " " & [MiddleInit] & ".";
