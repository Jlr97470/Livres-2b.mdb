CREATE TABLE [tblCustomers] (
  [CustomerID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [LastName] VARCHAR (50),
  [FirstName] VARCHAR (50),
  [MiddleInit] VARCHAR (1),
  [Address] VARCHAR (255),
  [City] VARCHAR (50),
  [StateOrProvince] VARCHAR (20),
  [PostalCode] VARCHAR (20),
  [Country] VARCHAR (50),
  [PhoneNumber] VARCHAR (30),
  [PayBy] SHORT,
  [CCNumber] VARCHAR (20),
  [EmailAddress] VARCHAR (50),
  [Password] VARCHAR (16),
  [Admin] BIT,
   CONSTRAINT 
)
