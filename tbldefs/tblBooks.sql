CREATE TABLE [tblBooks] (
  [ISBNNumber] VARCHAR (13) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Title] VARCHAR (75),
  [CoverPix] LONGBINARY,
  [CopyrightYear] SHORT,
  [SuggPrice] CURRENCY,
  [EditionNumber] LONG,
  [OutOfPrint] BIT,
  [Pages] SHORT,
  [Disk] VARCHAR (15),
  [LinerNotes] LONGTEXT,
  [TOCDoc] LONGTEXT
)
