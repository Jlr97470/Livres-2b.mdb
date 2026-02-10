SELECT
  tblStores.StoreID,
  tblStores.PostalCode,
  tblStores.StoreName,
  tblStores.Address,
  tblStores.City,
  tblStores.StateOrProvince,
  tblStores.Country,
  tblStores.PhoneNumber,
  tblStores.FaxNumber,
  tblStores.Email,
  tblStores.WebAddress,
  Left([PhoneNumber], 5) AS PhoneCode
FROM
  tblStores
ORDER BY
  [tblStores].[PostalCode],
  [tblStores].[StoreName];
