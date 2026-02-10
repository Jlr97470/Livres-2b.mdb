CREATE TABLE [tblOrders] (
  [OrderID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [CustomerID] LONG CONSTRAINT [{47D921B4-9F09-11D2-99E8-A76583F5D36C}] REFERENCES [tblCustomers] ([CustomerID]) ON DELETE CASCADE ,
  [StoreID] LONG CONSTRAINT [{47D921B5-9F09-11D2-99E8-A76583F5D36C}] REFERENCES [tblStores] ([StoreID]),
  [OrderDate] DATETIME,
  [ShipName] VARCHAR (50),
  [ShipAddress] VARCHAR (255),
  [ShipCity] VARCHAR (50),
  [ShipStateOrProvince] VARCHAR (20),
  [ShipPostalCode] VARCHAR (20),
  [ShipCountry] VARCHAR (50),
  [ShipPhoneNumber] VARCHAR (30),
  [PayBy] SHORT,
  [CCNumber] VARCHAR (20),
  [Printed] BIT
)
