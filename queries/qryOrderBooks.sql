SELECT
  tblOrderDetails.*,
  tblBooks.Title,
  tblBooks.EditionNumber,
  tblBooks.SuggPrice
FROM
  tblBooks
  INNER JOIN tblOrderDetails ON [tblBooks].[ISBNNumber] = [tblOrderDetails].[ISBNNumber];
