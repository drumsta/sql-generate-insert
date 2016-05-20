# INSERT statement(s) generator #
Generates INSERT statement(s) for data in a table.

## Purpose ##
- To regenerate data at another location.
- To script table or view data populated in automated way.
- To script setup data populated in automated/manual way.

## Installation ##

* Pre-requisites: MS SQL Server 2005 or later
* Download a copy of the `GenerateInsert.sql`
* Open SQL Server Management studio and load `GenerateInsert.sql`
* Select a database to install the stored procedure to
* Click Execute from the toolbar, this should run with a result of `Command Completely Successfully`

## Usage ##

### Quick example ###

```
USE [AdventureWorks];
GO
EXECUTE dbo.GenerateInsert @ObjectName = N'Person.AddressType';
```
This will generate the following script:
```
SET NOCOUNT ON
SET IDENTITY_INSERT Person.AddressType ON
INSERT INTO Person.AddressType
([AddressTypeID],[Name],[rowguid],[ModifiedDate])
VALUES
 (1,N'Billing','B84F78B1-4EFE-4A0E-8CB7-70E9F112F886',CONVERT(datetime,'2002-06-01 00:00:00.000',121))
,(2,N'Home','41BC2FF6-F0FC-475F-8EB9-CEC0805AA0F2',CONVERT(datetime,'2002-06-01 00:00:00.000',121))
,(3,N'Main Office','8EEEC28C-07A2-4FB9-AD0A-42D4A0BBC575',CONVERT(datetime,'2002-06-01 00:00:00.000',121))
,(4,N'Primary','24CB3088-4345-47C4-86C5-17B535133D1E',CONVERT(datetime,'2002-06-01 00:00:00.000',121))
,(5,N'Shipping','B29DA3F8-19A3-47DA-9DAA-15C84F4A83A5',CONVERT(datetime,'2002-06-01 00:00:00.000',121))
,(6,N'Archive','A67F238A-5BA2-444B-966C-0467ED9C427F',CONVERT(datetime,'2002-06-01 00:00:00.000',121))
SET IDENTITY_INSERT Person.AddressType OFF
```

### Example using SELECT syntax ###

```
EXECUTE dbo.GenerateInsert @ObjectName = N'Person.AddressType'
,@UseSelectSyntax=1
,@UseColumnAliasInSelect=1
,@GenerateOneColumnPerLine=1;
```
This will generate the following script:
```
SET NOCOUNT ON
SET IDENTITY_INSERT Person.AddressType ON
INSERT INTO Person.AddressType
([AddressTypeID]
,[Name]
,[rowguid]
,[ModifiedDate]
)
SELECT 1 [AddressTypeID]
,N'Billing' [Name]
,'B84F78B1-4EFE-4A0E-8CB7-70E9F112F886' [rowguid]
,CONVERT(datetime,'2002-06-01 00:00:00.000',121) [ModifiedDate]
UNION
SELECT 2 [AddressTypeID]
,N'Home' [Name]
,'41BC2FF6-F0FC-475F-8EB9-CEC0805AA0F2' [rowguid]
,CONVERT(datetime,'2002-06-01 00:00:00.000',121) [ModifiedDate]
UNION
SELECT 3 [AddressTypeID]
,N'Main Office' [Name]
,'8EEEC28C-07A2-4FB9-AD0A-42D4A0BBC575' [rowguid]
,CONVERT(datetime,'2002-06-01 00:00:00.000',121) [ModifiedDate]
UNION
SELECT 4 [AddressTypeID]
,N'Primary' [Name]
,'24CB3088-4345-47C4-86C5-17B535133D1E' [rowguid]
,CONVERT(datetime,'2002-06-01 00:00:00.000',121) [ModifiedDate]
UNION
SELECT 5 [AddressTypeID]
,N'Shipping' [Name]
,'B29DA3F8-19A3-47DA-9DAA-15C84F4A83A5' [rowguid]
,CONVERT(datetime,'2002-06-01 00:00:00.000',121) [ModifiedDate]
UNION
SELECT 6 [AddressTypeID]
,N'Archive' [Name]
,'A67F238A-5BA2-444B-966C-0467ED9C427F' [rowguid]
,CONVERT(datetime,'2002-06-01 00:00:00.000',121) [ModifiedDate]
SET IDENTITY_INSERT Person.AddressType OFF
```

### Select results into table variable for later reuse ###
The example below is pretty tricky because simple approach `INSERT INTO... EXECUTE dbo.GenerateInsert;` ends up with `INSERT EXEC statement cannot be nested.` Some pre-requisites are required in advance, i.e. ad hoc distributed queries should be allowed on the server, connection is made using Windows authentication.
```
DECLARE @Results table (TableRow varchar(max));
DECLARE @sql nvarchar(max) =
'SELECT * FROM OPENROWSET (
''SQLNCLI'',
''Server=(local);Database=' + DB_NAME() + ';Trusted_Connection=yes;'',
''EXECUTE dbo.GenerateInsert @ObjectName = N''''Person.AddressType''''
,@OmmitInsertColumnList=1
,@GenerateSingleInsertPerRow=1
,@FormatCode=0
,@GenerateGo=0
,@PrintGeneratedCode=0
;''
)';

INSERT INTO @Results
EXECUTE sp_executesql @sql;

SELECT *
FROM @Results;
```
This will generate the following script:
```
SET NOCOUNT ON
SET IDENTITY_INSERT Person.AddressType ON
INSERT INTO Person.AddressType  VALUES (1,N'Billing','B84F78B1-4EFE-4A0E-8CB7-70E9F112F886',CONVERT(datetime,'2002-06-01 00:00:00.000',121))
INSERT INTO Person.AddressType  VALUES (2,N'Home','41BC2FF6-F0FC-475F-8EB9-CEC0805AA0F2',CONVERT(datetime,'2002-06-01 00:00:00.000',121))
INSERT INTO Person.AddressType  VALUES (3,N'Main Office','8EEEC28C-07A2-4FB9-AD0A-42D4A0BBC575',CONVERT(datetime,'2002-06-01 00:00:00.000',121))
INSERT INTO Person.AddressType  VALUES (4,N'Primary','24CB3088-4345-47C4-86C5-17B535133D1E',CONVERT(datetime,'2002-06-01 00:00:00.000',121))
INSERT INTO Person.AddressType  VALUES (5,N'Shipping','B29DA3F8-19A3-47DA-9DAA-15C84F4A83A5',CONVERT(datetime,'2002-06-01 00:00:00.000',121))
INSERT INTO Person.AddressType  VALUES (6,N'Archive','A67F238A-5BA2-444B-966C-0467ED9C427F',CONVERT(datetime,'2002-06-01 00:00:00.000',121))
SET IDENTITY_INSERT Person.AddressType OFF
```

### Script all tables ###
```
DECLARE @Name nvarchar(261);
DECLARE TableCursor CURSOR LOCAL FAST_FORWARD FOR
SELECT QUOTENAME(s.name) + '.' + QUOTENAME(t.name) ObjectName
FROM sys.tables t
  INNER JOIN sys.schemas s ON s.schema_id = t.schema_id
WHERE t.name NOT LIKE 'sys%'
FOR READ ONLY
;
OPEN TableCursor;
FETCH NEXT FROM TableCursor INTO @Name;

WHILE @@FETCH_STATUS = 0
BEGIN
  EXECUTE dbo.GenerateInsert @ObjectName = @Name;

  FETCH NEXT FROM TableCursor INTO @Name;
END

CLOSE TableCursor;
DEALLOCATE TableCursor;
```
