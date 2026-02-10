SELECT
  NullLongID As StoreID,
  "<No specific store>" As StoreName,
  Null As CityState,
  Null As State
From
  ztblNullEntry
UNION
SELECT
  StoreID,
  StoreName,
  [City] & ", " & [StateOrProvince] AS CityState,
  StateOrProvince As State
FROM
  tblStores
ORDER BY
  State,
  CityState,
  StoreName;
