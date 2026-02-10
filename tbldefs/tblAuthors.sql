CREATE TABLE [tblAuthors] (
  [AuthorID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [FirstName] VARCHAR (50),
  [MiddleInit] VARCHAR (1),
  [LastName] VARCHAR (50),
  [EmailAddress] VARCHAR (50),
  [Photograph] LONGBINARY,
  [Bio] LONGTEXT
)
