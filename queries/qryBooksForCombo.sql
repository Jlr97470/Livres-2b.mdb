SELECT
  tblBooks.ISBNNumber,
  tblBooks.Title,
  tblBooks.EditionNumber
FROM
  tblBooks
ORDER BY
  [tblBooks].[Title],
  [tblBooks].[EditionNumber];
