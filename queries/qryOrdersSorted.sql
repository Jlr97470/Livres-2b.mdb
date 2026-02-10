SELECT
  tblCustomers.LastName,
  tblCustomers.FirstName,
  tblCustomers.MiddleInit,
  tblCustomers.Address,
  tblCustomers.City,
  tblCustomers.StateOrProvince,
  tblCustomers.PostalCode,
  tblCustomers.Country,
  tblCustomers.PhoneNumber,
  [tblCustomers].[PayBy] AS CustPayBy,
  [tblCustomers].[CCNumber] AS CustCCNumber,
  tblOrders.OrderID,
  tblOrders.CustomerID,
  tblOrders.StoreID,
  tblOrders.OrderDate,
  tblOrders.ShipName,
  tblOrders.ShipAddress,
  tblOrders.ShipCity,
  tblOrders.ShipStateOrProvince,
  tblOrders.ShipPostalCode,
  tblOrders.ShipCountry,
  tblOrders.ShipPhoneNumber,
  tblOrders.PayBy,
  tblOrders.CCNumber,
  tblOrders.Printed,
  tblStores.StoreName,
  [tblStores].[Address] AS StAddress,
  [tblStores].[City] AS StCity,
  [tblStores].[PostalCode] AS StPostalCode,
  [tblStores].[StateOrProvince] AS StStateOrProvince,
  [tblStores].[Country] AS StCountry
FROM
  tblStores
  RIGHT JOIN (
    tblCustomers
    INNER JOIN tblOrders ON [tblCustomers].[CustomerID] = [tblOrders].[CustomerID]
  ) ON [tblStores].[StoreID] = [tblOrders].[StoreID];
