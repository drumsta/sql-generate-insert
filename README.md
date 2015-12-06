# INSERT statement(s) generator #
Generates INSERT statement(s) for data in a table.

## Purpose ##
- To regenerate data at another location.
- To script table or view data populated in automated way.
- To script setup data populated in automated/manual way.

## Download and build instructions: ##

* Pre-requisites: MS SQL Server 2008 or later
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
```
DECLARE @Results table (TableRow varchar(max));
DECLARE @sql nvarchar(max) =
'SELECT * FROM OPENROWSET (
''SQLNCLI'',
''Server=(local);Database='+(SELECT DB_NAME())+';Trusted_Connection=yes;'',
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

## Arguments ##
```
  @ObjectName
    Format: [schema_name.]object_name
    Specifies the name of a table or view to generate the INSERT statement(s) for
  @TargetObjectName
    Specifies the name of target table or view to insert into
  @OmmitInsertColumnList
    When 0 then syntax is like INSERT INTO object (column_list)...
    When 1 then syntax is like INSERT INTO object...
  @GenerateSingleInsertPerRow bit = 0
    When 0 then only one INSERT statement is generated for all rows
    When 1 then separate INSERT statement is generated for every row
  @UseSelectSyntax bit = 0
    When 0 then syntax is like INSERT INTO object (column_list) VALUES(...)
    When 1 then syntax is like INSERT INTO object (column_list) SELECT...
  @UseColumnAliasInSelect bit = 0
    Has effect only when @UseSelectSyntax = 1
    When 0 then syntax is like SELECT 'value1','value2'
    When 1 then syntax is like SELECT 'value1' column1,'value2' column2
  @FormatCode bit = 1
    When 0 then no Line Feeds are generated
    When 1 then additional Line Feeds are generated for better readibility
  @GenerateOneColumnPerLine bit = 0
    When 0 then syntax is like SELECT 'value1','value2'...
      or VALUES('value1','value2')...
    When 1 then syntax is like
         SELECT
         'value1'
         ,'value2'
         ...
      or VALUES(
         'value1'
         ,'value2'
         )...
  @GenerateGo bit = 0
    When 0 then no GO commands are generated
    When 1 then GO commands are generated after each INSERT
  @PrintGeneratedCode bit = 1
    When 0 then generated code will be printed using PRINT command
    When 1 then generated code will be selected using SELECT statement 
  @TopExpression varchar(max) = NULL
    When supplied then INSERT statements are generated only for TOP rows
    Format: (expression) [PERCENT]
    Example: @TopExpression='(5)' is equivalent to SELECT TOP (5)
    Example: @TopExpression='(50) PERCENT' is equivalent to SELECT TOP (5) PERCENT
  @SearchCondition varchar(max) = NULL
    When supplied then specifies the search condition for the rows returned by the query
    Format: <search_condition>
    Example: @SearchCondition='column1 != ''test''' is equivalent to WHERE column1 != 'test'
  @OmmitUnsupportedDataTypes bit = 1
    When 0 then error is raised on unsupported data types
    When 1 then columns with unsupported data types are excluded from generation process
  @PopulateIdentityColumn bit = 1
    When 0 then identity columns are excluded from generation process
    When 1 then identity column values are preserved on insertion
  @PopulateTimestampColumn bit = 0
    When 0 then rowversion/timestamp column is inserted using DEFAULT value
    When 1 then rowversion/timestamp column values are preserved on insertion,
      useful when restoring into archive table as varbinary(8) to preserve history
  @PopulateComputedColumn bit = 0
    When 0 then computed columns are excluded from generation process
    When 1 then computed column values are preserved on insertion,
      useful when restoring into archive table as scalar values to preserve history
  @ShowWarnings bit = 1
    When 0 then no warnings are printed.
    When 1 then warnings are printed if columns with unsupported data types
      have been excluded from generation process
    Has effect only when @OmmitUnsupportedDataTypes = 1
  @Debug bit = 0
    When 0 then no debug information are printed.
    When 1 then constructed SQL statements are printed for later examination
```
