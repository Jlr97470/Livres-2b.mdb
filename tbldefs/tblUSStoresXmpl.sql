CREATE TABLE [tblUSStoresXmpl] (
  [StoreID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [StoreName] VARCHAR (50),
  [Address] VARCHAR (255),
  [City] VARCHAR (50),
  [PostalCode] VARCHAR (20),
  [StateOrProvince] VARCHAR (20),
  [Country] VARCHAR (50),
  [PhoneNumber] VARCHAR (15),
  [FaxNumber] VARCHAR (30),
  [Email] VARCHAR (100),
  [WebAddress] LONGTEXT
)
