SELECT
  tblCategories.Category,
  tblClasses.ClassName,
  tblCategories.CategoryID
FROM
  tblClasses
  INNER JOIN tblCategories ON [tblClasses].[ClassID] = [tblCategories].[ClassID]
ORDER BY
  [tblCategories].[Category],
  [tblClasses].[ClassName];
