SELECT
  tblAuthors.AuthorID,
  tblCategories.CategoryID
FROM
  tblCategories
  INNER JOIN (
    (
      tblBooks
      INNER JOIN (
        tblAuthors
        INNER JOIN tblBookAuthor ON [tblAuthors].[AuthorID] = [tblBookAuthor].[AuthorID]
      ) ON [tblBooks].[ISBNNumber] = [tblBookAuthor].[ISBNNumber]
    )
    INNER JOIN tblBookCategories ON [tblBooks].[ISBNNumber] = [tblBookCategories].[ISBNNumber]
  ) ON [tblCategories].[CategoryID] = [tblBookCategories].[CategoryID];
