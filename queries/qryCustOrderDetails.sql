SELECT
  tblCustomers.*,
  [LastName] & ", " & [FirstName] & " " & [MiddleInit] & "." AS CustName,
  tblOrders.OrderID,
  tblOrders.OrderDate,
  tblOrderDetails.ISBNNumber,
  tblOrderDetails.Quantity,
  tblOrderDetails.Discount,
  tblBooks.Title,
  tblBooks.SuggPrice
FROM
  (
    tblCustomers
    INNER JOIN tblOrders ON [tblCustomers].[CustomerID] = [tblOrders].[CustomerID]
  )
  INNER JOIN (
    tblBooks
    INNER JOIN tblOrderDetails ON [tblBooks].[ISBNNumber] = [tblOrderDetails].[ISBNNumber]
  ) ON [tblOrders].[OrderID] = [tblOrderDetails].[OrderID];
