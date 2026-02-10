SELECT
  tblStores.StoreID,
  tblStores.StoreName,
  tblStores.Address,
  tblStores.City,
  tblStores.StateOrProvince,
  tblStores.PostalCode,
  tblStores.Country,
  tblStores.PhoneNumber,
  tblStores.FaxNumber,
  tblStores.Email,
  tblStores.WebAddress,
  Left([PhoneNumber], 5) AS PhoneCode
FROM
  tblStores;
