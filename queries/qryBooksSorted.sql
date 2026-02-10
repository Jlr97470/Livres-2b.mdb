SELECT
  tblBooks.ISBNNumber,
  tblBooks.Title,
  tblBooks.CoverPix,
  tblBooks.CopyrightYear,
  tblBooks.SuggPrice,
  tblBooks.EditionNumber,
  tblBooks.OutOfPrint,
  tblBooks.Pages,
  tblBooks.Disk,
  tblBooks.LinerNotes,
  tblBooks.TOCDoc
FROM
  tblBooks
ORDER BY
  [tblBooks].[Title],
  [tblBooks].[EditionNumber];
