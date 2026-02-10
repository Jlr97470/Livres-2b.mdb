SELECT
  tblStores.StoreID,
  tblStores.StoreName,
  [City] & ", " & [StateOrProvince] AS CityState
FROM
  tblStores
ORDER BY
  [tblStores].[StateOrProvince],
  [tblStores].[City],
  [tblStores].[StoreName];
