IF EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'dbo.GenerateInsert') AND type in (N'P', N'PC'))
  DROP PROCEDURE dbo.GenerateInsert;
GO

CREATE PROCEDURE dbo.GenerateInsert
(
  @ObjectName nvarchar(261)
, @TargetObjectName nvarchar(261) = NULL
, @OmmitInsertColumnList bit = 0
, @GenerateSingleInsertPerRow bit = 0
, @UseSelectSyntax bit = 0
, @UseColumnAliasInSelect bit = 0
, @FormatCode bit = 1
, @GenerateOneColumnPerLine bit = 0
, @GenerateGo bit = 0
, @PrintGeneratedCode bit = 1
, @TopExpression varchar(max) = NULL
, @SearchCondition varchar(max) = NULL
, @OmmitUnsupportedDataTypes bit = 1
, @PopulateIdentityColumn bit = 1
, @PopulateTimestampColumn bit = 0
, @PopulateComputedColumn bit = 0
, @ShowWarnings bit = 1
, @Debug bit = 0
)
AS
/*******************************************************************************
Procedure: GenerateInsert (Build 1)
Decription: Generates INSERT statement(s) for data in a table.
Purpose: To regenerate data at another location.
  To script data populated in automated way.
  To script setup data populated in automated/manual way.
Project page: http://github.com/drumsta/sql-generate-insert

Arguments:
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
*******************************************************************************/
BEGIN
SET NOCOUNT ON;

DECLARE @CrLf char(2) = CHAR(13) + CHAR(10);
DECLARE @ColumnName sysname;
DECLARE @DataType sysname;
DECLARE @ColumnList nvarchar(max) = '';
DECLARE @SelectList nvarchar(max) = '';
DECLARE @SelectStatement nvarchar(max) = '';
DECLARE @OmmittedColumnList nvarchar(max) = '';
DECLARE @InsertSql varchar(max) = 'INSERT INTO ' + COALESCE(@TargetObjectName,@ObjectName);
DECLARE @ValuesSql varchar(max) = 'VALUES (';
DECLARE @SelectSql varchar(max) = 'SELECT ';
DECLARE @TableData table (TableRow varchar(max));
DECLARE @Results table (TableRow varchar(max));
DECLARE @TableRow nvarchar(max);
DECLARE @RowNo int;

IF PARSENAME(@ObjectName,3) IS NOT NULL
  OR PARSENAME(@ObjectName,4) IS NOT NULL
BEGIN
  RAISERROR('Server and database names are not allowed to specify in @ObjectName parameter. Required format is [schema_name.]object_name',16,1);
  RETURN -1;
END

IF OBJECT_ID(@ObjectName,N'U') IS NULL
  AND OBJECT_ID(@ObjectName,N'V') IS NULL
BEGIN
  RAISERROR(N'User table or view %s not found or insuficient permission to query the table or view.',16,1,@ObjectName);
  RETURN -1;
END

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = PARSENAME(@ObjectName,1) AND TABLE_TYPE IN ('BASE TABLE','VIEW') AND (TABLE_SCHEMA = PARSENAME(@ObjectName,2) OR PARSENAME(@ObjectName,2) IS NULL))
BEGIN
  RAISERROR(N'User table or view %s not found or insuficient permission to query the table or view.',16,1,@ObjectName);
  RETURN -1;
END

DECLARE ColumnCursor CURSOR LOCAL FAST_FORWARD FOR
SELECT c.name ColumnName
,TYPE_NAME(c.user_type_id) DataType
FROM sys.objects o
  INNER JOIN sys.columns c ON c.object_id = o.object_id
WHERE o.type IN (N'U',N'V') -- USER_TABLE,VIEW
  AND (o.object_id = OBJECT_ID(@ObjectName)
    OR o.name = @ObjectName)
  AND (COLUMNPROPERTY(c.object_id,c.name,'IsIdentity') != 1
    OR @PopulateIdentityColumn = 1)
  AND (COLUMNPROPERTY(c.object_id,c.name,'IsComputed') != 1
    OR @PopulateComputedColumn = 1)
ORDER BY COLUMNPROPERTY(c.object_id,c.name,'ordinal') -- ORDINAL_POSITION
FOR READ ONLY
;
OPEN ColumnCursor;
FETCH NEXT FROM ColumnCursor INTO @ColumnName,@DataType;

WHILE @@FETCH_STATUS = 0
BEGIN
  -- Handle different data types
  DECLARE @ColumnExpression varchar(max);
  SET @ColumnExpression = 
    CASE
    WHEN @DataType IN ('char','varchar','text','uniqueidentifier')
    THEN 'ISNULL(''''''''+REPLACE(CONVERT(varchar(max),'+  QUOTENAME(@ColumnName) + '),'''''''','''''''''''')+'''''''',''NULL'') COLLATE database_default'
      
    WHEN @DataType IN ('nchar','nvarchar','sysname','ntext','sql_variant','xml')
    THEN 'ISNULL(''N''''''+REPLACE(CONVERT(nvarchar(max),'+  QUOTENAME(@ColumnName) + '),'''''''','''''''''''')+'''''''',''NULL'') COLLATE database_default'
      
    WHEN @DataType IN ('int','bigint','smallint','tinyint','decimal','numeric','bit')
    THEN 'ISNULL(CONVERT(varchar(max),'+  QUOTENAME(@ColumnName) + '),''NULL'') COLLATE database_default'
      
    WHEN @DataType IN ('float','real','money','smallmoney')
    THEN 'ISNULL(CONVERT(varchar(max),'+  QUOTENAME(@ColumnName) + ',2),''NULL'') COLLATE database_default'
      
    WHEN @DataType IN ('datetime','smalldatetime','date','time','datetime2','datetimeoffset')
    THEN '''CONVERT('+@DataType+',''+ISNULL(''''''''+CONVERT(varchar(max),'+  QUOTENAME(@ColumnName) + ',121)+'''''''',''NULL'') COLLATE database_default' + '+'',121)'''

    WHEN @DataType IN ('rowversion','timestamp')
    THEN
      CASE WHEN @PopulateTimestampColumn = 1
      THEN '''CONVERT(varbinary(max),''+ISNULL(''''''''+CONVERT(varchar(max),CONVERT(varbinary(max),'+  QUOTENAME(@ColumnName) + '),1)+'''''''',''NULL'') COLLATE database_default' + '+'',1)'''
      ELSE '''NULL''' END

    WHEN @DataType IN ('binary','varbinary','image')
    THEN '''CONVERT(varbinary(max),''+ISNULL(''''''''+CONVERT(varchar(max),CONVERT(varbinary(max),'+  QUOTENAME(@ColumnName) + '),1)+'''''''',''NULL'') COLLATE database_default' + '+'',1)'''

    WHEN @DataType IN ('geography')
    -- convert geography to text: ?? column.STAsText();
    -- convert text to geography: ?? geography::STGeomFromText('LINESTRING(-122.360 47.656, -122.343 47.656 )', 4326);
    THEN NULL

    ELSE NULL END;

  IF @ColumnExpression IS NULL
    AND @OmmitUnsupportedDataTypes != 1
  BEGIN
    RAISERROR(N'Datatype %s is not supported. Use @OmmitUnsupportedDataTypes to exclude unsupported columns.',16,1,@DataType);
    RETURN -1;
  END

  IF @ColumnExpression IS NULL
  BEGIN
    SET @OmmittedColumnList = @OmmittedColumnList
      + CASE WHEN @OmmittedColumnList != '' THEN ', ' ELSE '' END
      + QUOTENAME(@ColumnName)
      + ' ' + @DataType;
  END

  IF @ColumnExpression IS NOT NULL
  BEGIN
    SET @ColumnList = @ColumnList
      + CASE WHEN @ColumnList != '' THEN ',' ELSE '' END
      + QUOTENAME(@ColumnName)
      + CASE WHEN @GenerateOneColumnPerLine = 1 THEN @CrLf ELSE '' END;
  
    SET @SelectList = @SelectList
      + CASE WHEN @SelectList != '' THEN '+'',''+' + @CrLf ELSE '' END
      + @ColumnExpression
      + CASE WHEN @UseColumnAliasInSelect = 1 AND @UseSelectSyntax = 1 THEN '+'' ' + QUOTENAME(@ColumnName) + '''' ELSE '' END
      + CASE WHEN @GenerateOneColumnPerLine = 1 THEN '+CHAR(13)+CHAR(10)' ELSE '' END;
  END

  FETCH NEXT FROM ColumnCursor INTO @ColumnName,@DataType;
END

CLOSE ColumnCursor;
DEALLOCATE ColumnCursor;

IF NULLIF(@ColumnList,'') IS NULL
BEGIN
  RAISERROR(N'No columns to select.',16,1);
  RETURN -1;
END

IF @Debug = 1
BEGIN
  PRINT '--Column list';
  PRINT @ColumnList;
END

IF NULLIF(@OmmittedColumnList,'') IS NOT NULL
  AND @ShowWarnings = 1
BEGIN
  PRINT(N'--WARNING: The following columns have been ommitted because of unsupported datatypes: ' + @OmmittedColumnList);
END

IF @GenerateSingleInsertPerRow = 1
BEGIN
  SET @SelectList = 
    '''' + @InsertSql + '''+' + @CrLf
    + CASE WHEN @FormatCode = 1
      THEN 'CHAR(13)+CHAR(10)+' + @CrLf
      ELSE ''' ''+'
      END
    + CASE WHEN @OmmitInsertColumnList = 1
      THEN ''
      ELSE '''(' + @ColumnList + ')''+' + @CrLf
      END
    + CASE WHEN @FormatCode = 1
      THEN 'CHAR(13)+CHAR(10)+' + @CrLf
      ELSE ''' ''+'
      END
    + CASE WHEN @UseSelectSyntax = 1
      THEN '''' + @SelectSql + '''+'
      ELSE '''' + @ValuesSql + '''+'
      END
    + @CrLf
    + @SelectList
    + CASE WHEN @UseSelectSyntax = 1
      THEN ''
      ELSE '+' + @CrLf + ''')'''
      END
    + CASE WHEN @GenerateGo = 1
      THEN '+' + @CrLf + 'CHAR(13)+CHAR(10)+' + @CrLf + '''GO'''
      ELSE ''
      END
  ;
END ELSE BEGIN
  SET @SelectList =
    CASE WHEN @UseSelectSyntax = 1
      THEN '''' + @SelectSql + '''+'
      ELSE '''(''+'
      END
    + @CrLf
    + @SelectList
    + CASE WHEN @UseSelectSyntax = 1
      THEN ''
      ELSE '+' + @CrLf + ''')'''
      END
  ;
END

SET @SelectStatement = 'SELECT'
  + CASE WHEN NULLIF(@TopExpression,'') IS NOT NULL
    THEN ' TOP ' + @TopExpression
    ELSE '' END
  + @CrLf + @SelectList + @CrLf
  + 'FROM ' + @ObjectName
  + CASE WHEN NULLIF(@SearchCondition,'') IS NOT NULL
    THEN @CrLf + 'WHERE ' + @SearchCondition
    ELSE '' END
;

IF @Debug = 1
BEGIN
  PRINT '--Select statement';
  PRINT @SelectStatement;
END

INSERT INTO @TableData
EXECUTE (@SelectStatement);

INSERT INTO @Results
SELECT '--INSERTs generated by GenerateInsert (Build 1)'
UNION SELECT '--Project page: http://github.com/drumsta/sql-generate-insert'
UNION SELECT 'SET NOCOUNT ON'

IF @PopulateIdentityColumn = 1
BEGIN
  INSERT INTO @Results
  SELECT 'SET IDENTITY_INSERT ' + COALESCE(@TargetObjectName,@ObjectName) + ' ON'
END

IF @GenerateSingleInsertPerRow = 1
BEGIN
  INSERT INTO @Results
  SELECT TableRow
  FROM @TableData
END ELSE BEGIN
  IF @FormatCode = 1
  BEGIN
    INSERT INTO @Results
    SELECT @InsertSql;

    IF @OmmitInsertColumnList != 1
    BEGIN
      INSERT INTO @Results
      SELECT '(' + @ColumnList + ')';
    END

    IF @UseSelectSyntax != 1
    BEGIN
      INSERT INTO @Results
      SELECT 'VALUES';
    END
  END ELSE BEGIN
    INSERT INTO @Results
    SELECT @InsertSql
      + CASE WHEN @OmmitInsertColumnList = 1 THEN '' ELSE ' (' + @ColumnList + ')' END
      + CASE WHEN @UseSelectSyntax = 1 THEN '' ELSE ' VALUES' END
  END

  SET @RowNo = 0;
  DECLARE DataCursor CURSOR LOCAL FAST_FORWARD FOR
  SELECT TableRow
  FROM @TableData
  FOR READ ONLY
  ;
  OPEN DataCursor;
  FETCH NEXT FROM DataCursor INTO @TableRow;

  WHILE @@FETCH_STATUS = 0
  BEGIN
    SET @RowNo = @RowNo + 1;

    INSERT INTO @Results
    SELECT
      CASE WHEN @UseSelectSyntax = 1
      THEN CASE WHEN @RowNo > 1 THEN 'UNION' + CASE WHEN @FormatCode = 1 THEN @CrLf ELSE ' ' END ELSE '' END
      ELSE CASE WHEN @RowNo > 1 THEN ',' ELSE ' ' END END
      + @TableRow;

    FETCH NEXT FROM DataCursor INTO @TableRow;
  END

  CLOSE DataCursor;
  DEALLOCATE DataCursor;

  IF @GenerateGo = 1
  BEGIN
    INSERT INTO @Results
    SELECT 'GO';
  END
END

IF @PopulateIdentityColumn = 1
BEGIN
  INSERT INTO @Results
  SELECT 'SET IDENTITY_INSERT ' + COALESCE(@TargetObjectName,@ObjectName) + ' OFF'
END

IF @PrintGeneratedCode = 1
BEGIN
  DECLARE ResultsCursor CURSOR LOCAL FAST_FORWARD FOR
  SELECT TableRow
  FROM @Results
  FOR READ ONLY
  ;
  OPEN ResultsCursor;
  FETCH NEXT FROM ResultsCursor INTO @TableRow;

  WHILE @@FETCH_STATUS = 0
  BEGIN
    PRINT(@TableRow);

    FETCH NEXT FROM ResultsCursor INTO @TableRow;
  END

  CLOSE ResultsCursor;
  DEALLOCATE ResultsCursor;
END ELSE BEGIN
  SELECT *
  FROM @Results;
END

END
GO
