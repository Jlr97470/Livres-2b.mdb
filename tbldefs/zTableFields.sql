CREATE TABLE [zTableFields] (
  [TableName] VARCHAR (50),
  [Sequence] LONG,
  [FieldName] VARCHAR (50),
  [Description] VARCHAR (255),
  [Type] VARCHAR (50),
  [Length] VARCHAR (50),
  [PKey] BIT,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([TableName], [Sequence])
)
