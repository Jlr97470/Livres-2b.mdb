SELECT
  tblBookCategories.ISBNNumber,
  tblBookCategories.CategoryID,
  tblClasses.ClassName,
  tblCategories.Category
FROM
  tblClasses
  INNER JOIN (
    tblCategories
    INNER JOIN tblBookCategories ON [tblCategories].[CategoryID] = [tblBookCategories].[CategoryID]
  ) ON [tblClasses].[ClassID] = [tblCategories].[ClassID];
