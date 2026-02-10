SELECT
  tblAuthors.AuthorID,
  tblAuthors.LastName,
  tblAuthors.FirstName,
  tblAuthors.MiddleInit,
  tblAuthors.EmailAddress,
  tblAuthors.Photograph,
  tblAuthors.Bio
FROM
  tblAuthors
ORDER BY
  [tblAuthors].[LastName],
  [tblAuthors].[FirstName];
