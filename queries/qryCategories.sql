SELECT
  tblcategories.CategoryID,
  tblcategories.ClassID,
  tblcategories.Category,
  tblClasses.ClassName
FROM
  tblClasses
  INNER JOIN tblcategories ON [tblClasses].[ClassID] = [tblcategories].[ClassID];
