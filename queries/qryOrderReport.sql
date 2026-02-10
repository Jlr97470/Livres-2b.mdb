SELECT
  qryOrdersSorted.LastName,
  qryOrdersSorted.FirstName,
  qryOrdersSorted.MiddleInit,
  qryOrdersSorted.Address,
  qryOrdersSorted.City,
  qryOrdersSorted.StateOrProvince,
  qryOrdersSorted.PostalCode,
  qryOrdersSorted.Country,
  qryOrdersSorted.PhoneNumber,
  [qryOrdersSorted].[PayBy] AS CustPayBy,
  [qryOrdersSorted].[CCNumber] AS CustCCNumber,
  qryOrdersSorted.OrderID,
  qryOrdersSorted.CustomerID,
  qryOrdersSorted.StoreID,
  qryOrdersSorted.OrderDate,
  qryOrdersSorted.ShipName,
  qryOrdersSorted.ShipAddress,
  qryOrdersSorted.ShipCity,
  qryOrdersSorted.ShipStateOrProvince,
  qryOrdersSorted.ShipPostalCode,
  qryOrdersSorted.ShipCountry,
  qryOrdersSorted.ShipPhoneNumber,
  qryOrdersSorted.PayBy,
  qryOrdersSorted.CCNumber,
  qryOrdersSorted.Printed,
  qryOrdersSorted.StoreName,
  qryOrdersSorted.StAddress,
  qryOrdersSorted.StCity,
  qryOrdersSorted.StPostalCode,
  qryOrdersSorted.StStateOrProvince,
  qryOrdersSorted.StCountry,
  tblOrderDetails.ISBNNumber,
  tblBooks.Title,
  tblBooks.EditionNumber,
  tblOrderDetails.Quantity,
  tblBooks.SuggPrice,
  tblOrderDetails.Discount
FROM
  tblBooks
  INNER JOIN (
    qryOrdersSorted
    INNER JOIN tblOrderDetails ON [qryOrdersSorted].[OrderID] = [tblOrderDetails].[OrderID]
  ) ON [tblBooks].[ISBNNumber] = [tblOrderDetails].[ISBNNumber]
ORDER BY
  [qryOrdersSorted].[LastName],
  [qryOrdersSorted].[FirstName],
  [qryOrdersSorted].[OrderDate];
