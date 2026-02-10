SELECT
  tblBookAuthor.ISBNNumber,
  tblBookAuthor.AuthorID,
  tblBookAuthor.SeqNo,
  tblBooks.Title,
  tblBooks.EditionNumber
FROM
  tblBooks
  INNER JOIN tblBookAuthor ON [tblBooks].[ISBNNumber] = [tblBookAuthor].[ISBNNumber];
