SELECT
  tblBookAuthor.ISBNNumber,
  tblBookAuthor.AuthorID,
  tblBookAuthor.SeqNo,
  tblAuthors.FirstName,
  tblAuthors.LastName,
  tblAuthors.MiddleInit
FROM
  tblAuthors
  INNER JOIN tblBookAuthor ON [tblAuthors].[AuthorID] = [tblBookAuthor].[AuthorID];
