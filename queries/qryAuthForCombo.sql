SELECT
  tblAuthors.AuthorID,
  [LastName] & ", " & [FirstName] & " " & [MiddleInit] AS AName
FROM
  tblAuthors
ORDER BY
  [LastName] & ", " & [FirstName] & " " & [MiddleInit];
