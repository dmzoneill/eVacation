--Microsoft SQL Server - Scripting
--Author: Peter Cejer   
--Created: 12th March 2002
--Last Updated: 13th March by Gemma Thompsom

/** Ensure that the scripts are run under the correct database!

Run steps in the following order:
    Run Step 1. Save Foreign Keys
    Run Step 2. Delete Foreign Keys
    Run the DTS package
    Run Step 2. Re-add Foreign Keys 
**/

-- Step 1. Save Foriegn Keys --
Use dbn_evacation
Go

IF EXISTS (SELECT 1 FROM sysobjects WHERE name = 'fkeys' AND type = 'U')
    drop table fkeys
GO

CREATE TABLE fkeys (PKTABLE_NAME varchar(32), FKTABLE_NAME varchar(32),
            FKCOLUMN1_NAME varchar(32) NULL,
            FKCOLUMN2_NAME varchar(32) NULL,
            FKCOLUMN3_NAME varchar(32) NULL,
            FKCOLUMN4_NAME varchar(32) NULL,
            FKCOLUMN5_NAME varchar(32) NULL,
            FKCOLUMN6_NAME varchar(32) NULL,
            FKCOLUMN7_NAME varchar(32) NULL,
            FKCOLUMN8_NAME varchar(32) NULL,
            FKCOLUMN9_NAME varchar(32) NULL,
            FKCOLUMN10_NAME varchar(32) NULL,
            FKCOLUMN11_NAME varchar(32) NULL,
            FKCOLUMN12_NAME varchar(32) NULL,
            FKCOLUMN13_NAME varchar(32) NULL,
            FKCOLUMN14_NAME varchar(32) NULL,
            FKCOLUMN15_NAME varchar(32) NULL,
            FKCOLUMN16_NAME varchar(32) NULL,
            FK_NAME varchar(32))


GO

/********************** GET FOREIGN KEYS ************************/


    DECLARE @pktable_id         int
    DECLARE @fktable_id         int
    DECLARE @fkfull_table_name  char(70)
    declare @order_by_pk        int




    create table #fkeys(
             pkdb_id        int NOT NULL,
             pktable_id     int NOT NULL,
             pkcolid        int NOT NULL,
             fkdb_id        int NOT NULL,
             fktable_id     int NOT NULL,
             fkcolid        int NOT NULL,
             KEY_SEQ        smallint NOT NULL,
             fk_id          int NOT NULL,
             pk_id          int NOT NULL)

    /*  SQL Server supports upto 16 PK/FK relationships between 2 tables */
    /*  Process syskeys for each relationship */
    /*  The inserts below adds a row to the temp table for each of the
        16 possible relationships */
    insert into #fkeys
        select
            r.rkeydbid,
            r.rkeyid,
            r.rkey1,
            r.fkeydbid,
            r.fkeyid,
            r.fkey1,
            1,
            r.constid,
            s.constid
        from
            sysreferences r, sysconstraints s
        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1
      union all
        select
            r.rkeydbid,
            r.rkeyid,
            r.rkey2,
            r.fkeydbid,
            r.fkeyid,
            r.fkey2,
            2,
            r.constid,
            s.constid
        from

            sysreferences r, sysconstraints s
        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1
      union all
        select
            r.rkeydbid,
            r.rkeyid,
            r.rkey3,
            r.fkeydbid,
            r.fkeyid,
            r.fkey3,
            3,
            r.constid,
            s.constid
        from
            sysreferences r, sysconstraints s
        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1
      union all
        select
            r.rkeydbid,
            r.rkeyid,
            r.rkey4,
            r.fkeydbid,
            r.fkeyid,
            r.fkey4,
            4,
            r.constid,
            s.constid
        from

            sysreferences r, sysconstraints s
        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1
      union all
        select
            r.rkeydbid,
            r.rkeyid,
            r.rkey5,
            r.fkeydbid,
            r.fkeyid,
            r.fkey5,
            5,
            r.constid,
            s.constid
        from
            sysreferences r, sysconstraints s
        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1
      union all
        select
            r.rkeydbid,
            r.rkeyid,
            r.rkey6,
            r.fkeydbid,
            r.fkeyid,

            r.fkey6,
            6,
            r.constid,
            s.constid
        from
            sysreferences r, sysconstraints s
        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1
      union all
        select
            r.rkeydbid,
            r.rkeyid,
            r.rkey7,

            r.fkeydbid,
            r.fkeyid,
            r.fkey7,
            7,
            r.constid,
            s.constid
        from
            sysreferences r, sysconstraints s
        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1
      union all
        select
            r.rkeydbid,
            r.rkeyid,

            r.rkey8,
            r.fkeydbid,
            r.fkeyid,
            r.fkey8,
            8,
            r.constid,
            s.constid
        from
            sysreferences r, sysconstraints s
        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1
      union all
        select
            r.rkeydbid,
            r.rkeyid,
            r.rkey9,
            r.fkeydbid,
            r.fkeyid,
            r.fkey9,
            9,
            r.constid,
            s.constid
        from
            sysreferences r, sysconstraints s
        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1
      union all
        select
            r.rkeydbid,
            r.rkeyid,
            r.rkey10,
            r.fkeydbid,
            r.fkeyid,
            r.fkey10,
            10,
            r.constid,
            s.constid
        from
            sysreferences r, sysconstraints s
        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1
      union all
        select
            r.rkeydbid,
            r.rkeyid,
            r.rkey11,
            r.fkeydbid,
            r.fkeyid,
            r.fkey11,
            11,
            r.constid,
            s.constid
        from
            sysreferences r, sysconstraints s
        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1
      union all
        select
            r.rkeydbid,

            r.rkeyid,
            r.rkey12,
            r.fkeydbid,
            r.fkeyid,
            r.fkey12,
            12,
            r.constid,
            s.constid
        from
            sysreferences r, sysconstraints s
        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1
      union all
        select
            r.rkeydbid,
            r.rkeyid,
            r.rkey13,
            r.fkeydbid,
            r.fkeyid,
            r.fkey13,
            13,
            r.constid,
            s.constid
        from
            sysreferences r, sysconstraints s
        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1
      union all
        select
            r.rkeydbid,
            r.rkeyid,
            r.rkey14,
            r.fkeydbid,
            r.fkeyid,
            r.fkey14,
            14,
            r.constid,
            s.constid
        from
            sysreferences r, sysconstraints s
        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1
      union all
        select
            r.rkeydbid,
            r.rkeyid,
            r.rkey15,
            r.fkeydbid,
            r.fkeyid,
            r.fkey15,
            15,
            r.constid,
            s.constid
        from
            sysreferences r, sysconstraints s

        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1
      union all
        select
            r.rkeydbid,
            r.rkeyid,
            r.rkey16,
            r.fkeydbid,
            r.fkeyid,

            r.fkey16,
            16,
            r.constid,
            s.constid
        from
            sysreferences r, sysconstraints s
        where   r.rkeyid = s.id
            AND (s.status & 0xf) = 1


        CREATE TABLE #foreignkeys (PKTABLE_NAME varchar(32), FKTABLE_NAME varchar(32),
            FKCOLUMN_NAME varchar(32), KEY_SEQ smallint, FK_NAME varchar(32))

        insert into #foreignkeys
        select
            PKTABLE_NAME = convert(varchar(32),o1.name),
            FKTABLE_NAME = convert(varchar(32),o2.name),
            FKCOLUMN_NAME = convert(varchar(32),c2.name),
            KEY_SEQ,
            FK_NAME = convert(varchar(128),OBJECT_NAME(fk_id))
        from #fkeys f,

            sysobjects o1, sysobjects o2,
            syscolumns c1, syscolumns c2
        where   o1.id = f.pktable_id
            AND o2.id = f.fktable_id
            AND c1.id = f.pktable_id
            AND c2.id = f.fktable_id
            AND c1.colid = f.pkcolid
            AND c2.colid = f.fkcolid

        CREATE TABLE #fgnkeys (PKTABLE_NAME varchar(32), FKTABLE_NAME varchar(32),
            FKCOLUMN1_NAME varchar(32) NULL,
            FKCOLUMN2_NAME varchar(32) NULL,
            FKCOLUMN3_NAME varchar(32) NULL,
            FKCOLUMN4_NAME varchar(32) NULL,
            FKCOLUMN5_NAME varchar(32) NULL,
            FKCOLUMN6_NAME varchar(32) NULL,
            FKCOLUMN7_NAME varchar(32) NULL,
            FKCOLUMN8_NAME varchar(32) NULL,
            FKCOLUMN9_NAME varchar(32) NULL,
            FKCOLUMN10_NAME varchar(32) NULL,
            FKCOLUMN11_NAME varchar(32) NULL,
            FKCOLUMN12_NAME varchar(32) NULL,
            FKCOLUMN13_NAME varchar(32) NULL,
            FKCOLUMN14_NAME varchar(32) NULL,
            FKCOLUMN15_NAME varchar(32) NULL,
            FKCOLUMN16_NAME varchar(32) NULL,
            FK_NAME varchar(32))

        INSERT INTO #fgnkeys
        SELECT  PKTABLE_NAME,
            FKTABLE_NAME,
            CASE 
             WHEN KEY_SEQ = 1 THEN FKCOLUMN_NAME
            END,
            CASE 
             WHEN KEY_SEQ = 2 THEN FKCOLUMN_NAME
            END,
            CASE 
             WHEN KEY_SEQ = 3 THEN FKCOLUMN_NAME
            END,
            CASE 
             WHEN KEY_SEQ = 4 THEN FKCOLUMN_NAME
            END,
            CASE 
             WHEN KEY_SEQ = 5 THEN FKCOLUMN_NAME
            END,
            CASE 
             WHEN KEY_SEQ = 6 THEN FKCOLUMN_NAME
            END,
            CASE 
             WHEN KEY_SEQ = 7 THEN FKCOLUMN_NAME
            END,
            CASE 
             WHEN KEY_SEQ = 8 THEN FKCOLUMN_NAME
            END,
            CASE 
             WHEN KEY_SEQ = 9 THEN FKCOLUMN_NAME
            END,
            CASE 
             WHEN KEY_SEQ = 10 THEN FKCOLUMN_NAME
            END,
            CASE 
             WHEN KEY_SEQ = 11 THEN FKCOLUMN_NAME
            END,
            CASE 
             WHEN KEY_SEQ = 12 THEN FKCOLUMN_NAME
            END,
            CASE 
             WHEN KEY_SEQ = 13 THEN FKCOLUMN_NAME
            END,
            CASE 
             WHEN KEY_SEQ = 14 THEN FKCOLUMN_NAME
            END,
            CASE 
             WHEN KEY_SEQ = 15 THEN FKCOLUMN_NAME
            END,
            CASE 
             WHEN KEY_SEQ = 16 THEN FKCOLUMN_NAME
            END,
            FK_NAME
        FROM #foreignkeys


        DELETE fkeys

        INSERT INTO fkeys
        SELECT  PKTABLE_NAME,
            FKTABLE_NAME,
            max(FKCOLUMN1_NAME),
            max(FKCOLUMN2_NAME),
            max(FKCOLUMN3_NAME),
            max(FKCOLUMN4_NAME),
            max(FKCOLUMN5_NAME),
            max(FKCOLUMN6_NAME),
            max(FKCOLUMN7_NAME),
            max(FKCOLUMN8_NAME),
            max(FKCOLUMN9_NAME),
            max(FKCOLUMN10_NAME),
            max(FKCOLUMN11_NAME),
            max(FKCOLUMN12_NAME),
            max(FKCOLUMN13_NAME),
            max(FKCOLUMN14_NAME),
            max(FKCOLUMN15_NAME),
            max(FKCOLUMN16_NAME),
            FK_NAME
        FROM #fgnkeys
        GROUP BY PKTABLE_NAME,
            FKTABLE_NAME,
            FK_NAME

GO


-- Step 2. Delete Foreign Keys --

/************* DELETE ALL FOREIGN KEYS ********************/
Use dbn_evacation
Go

DECLARE @fktable_name varchar(32),
    @fk_name varchar(32),
    @pktable_name varchar(32),
    @cmd varchar(255)


DECLARE fkeys_cursor 
CURSOR FOR 
SELECT PKTABLE_NAME, FKTABLE_NAME, FK_NAME FROM fkeys

OPEN fkeys_cursor

FETCH NEXT FROM fkeys_cursor INTO @pktable_name, @fktable_name, @fk_name

WHILE @@FETCH_STATUS > -1
BEGIN
    SELECT @cmd = 'ALTER TABLE ' + @fktable_name + ' DROP constraint ' + @fk_name 
    EXEC (@cmd)
    FETCH NEXT FROM fkeys_cursor INTO @pktable_name, @fktable_name, @fk_name
END


CLOSE fkeys_cursor
DEALLOCATE fkeys_cursor


GO


/*********************************************** 

    DTS THE DATA INTO THE DATABASE

************************************************/


-- Step 3. Re-ad Foreign Keys

/**************** ADD ALL FOREIGN KEYS BACK IN **********************/

DECLARE @fktable_name varchar(32),
    @fk_name varchar(32),
    @fkcolumn1_name varchar(32), 
    @fkcolumn2_name varchar(32), 
    @fkcolumn3_name varchar(32), 
    @fkcolumn4_name varchar(32), 

    @fkcolumn5_name varchar(32), 
    @fkcolumn6_name varchar(32), 
    @fkcolumn7_name varchar(32), 
    @fkcolumn8_name varchar(32), 
    @fkcolumn9_name varchar(32), 
    @fkcolumn10_name varchar(32), 
    @fkcolumn11_name varchar(32), 
    @fkcolumn12_name varchar(32), 
    @fkcolumn13_name varchar(32), 
    @fkcolumn14_name varchar(32), 
    @fkcolumn15_name varchar(32), 
    @fkcolumn16_name varchar(32), 
    @pktable_name varchar(32),
    @cmd varchar(255),
    @fkcolumns varchar (255)


DECLARE fkeys_cursor 
CURSOR FOR 
SELECT FKTABLE_NAME, FK_NAME, 
    FKCOLUMN1_NAME, 
    FKCOLUMN2_NAME, 
    FKCOLUMN3_NAME, 
    FKCOLUMN4_NAME, 
    FKCOLUMN5_NAME, 
    FKCOLUMN6_NAME, 
    FKCOLUMN7_NAME, 
    FKCOLUMN8_NAME, 
    FKCOLUMN9_NAME, 
    FKCOLUMN10_NAME, 
    FKCOLUMN11_NAME, 
    FKCOLUMN12_NAME, 
    FKCOLUMN13_NAME, 
    FKCOLUMN14_NAME, 
    FKCOLUMN15_NAME, 
    FKCOLUMN16_NAME, 
    PKTABLE_NAME FROM fkeys

OPEN fkeys_cursor

    FETCH NEXT FROM fkeys_cursor INTO @fktable_name, @fk_name, 
    @fkcolumn1_name, 
    @fkcolumn2_name, 
    @fkcolumn3_name, 
    @fkcolumn4_name, 
    @fkcolumn5_name, 
    @fkcolumn6_name, 
    @fkcolumn7_name, 
    @fkcolumn8_name, 
    @fkcolumn9_name, 
    @fkcolumn10_name, 
    @fkcolumn11_name, 
    @fkcolumn12_name, 
    @fkcolumn13_name, 
    @fkcolumn14_name, 
    @fkcolumn15_name, 
    @fkcolumn16_name, 
    @pktable_name

WHILE @@FETCH_STATUS > -1
BEGIN 

        SELECT @fkcolumns = @fkcolumn1_name
        IF  @fkcolumn2_name IS NOT NULL
            SELECT @fkcolumns = @fkcolumns + ',' + @fkcolumn2_name
        IF  @fkcolumn3_name IS NOT NULL
            SELECT @fkcolumns = @fkcolumns + ',' + @fkcolumn3_name
        IF  @fkcolumn4_name IS NOT NULL
            SELECT @fkcolumns = @fkcolumns + ',' + @fkcolumn4_name
        IF  @fkcolumn5_name IS NOT NULL
            SELECT @fkcolumns = @fkcolumns + ',' + @fkcolumn5_name
        IF  @fkcolumn6_name IS NOT NULL
            SELECT @fkcolumns = @fkcolumns + ',' + @fkcolumn6_name
        IF  @fkcolumn7_name IS NOT NULL
            SELECT @fkcolumns = @fkcolumns + ',' + @fkcolumn7_name
        IF  @fkcolumn8_name IS NOT NULL
            SELECT @fkcolumns = @fkcolumns + ',' + @fkcolumn8_name
        IF  @fkcolumn9_name IS NOT NULL
            SELECT @fkcolumns = @fkcolumns + ',' + @fkcolumn9_name
        IF  @fkcolumn10_name IS NOT NULL
            SELECT @fkcolumns = @fkcolumns + ',' + @fkcolumn10_name
        IF  @fkcolumn11_name IS NOT NULL
            SELECT @fkcolumns = @fkcolumns + ',' + @fkcolumn11_name
        IF  @fkcolumn12_name IS NOT NULL
            SELECT @fkcolumns = @fkcolumns + ',' + @fkcolumn12_name
        IF  @fkcolumn13_name IS NOT NULL
            SELECT @fkcolumns = @fkcolumns + ',' + @fkcolumn13_name
        IF  @fkcolumn14_name IS NOT NULL
            SELECT @fkcolumns = @fkcolumns + ',' + @fkcolumn14_name
        IF  @fkcolumn15_name IS NOT NULL
            SELECT @fkcolumns = @fkcolumns + ',' + @fkcolumn15_name
        IF  @fkcolumn16_name IS NOT NULL
            SELECT @fkcolumns = @fkcolumns + ',' + @fkcolumn16_name


        SELECT @cmd = 'ALTER TABLE ' + @fktable_name + ' ADD constraint ' + @fk_name 

            + ' FOREIGN KEY (' + @fkcolumns + ') REFERENCES ' + @pktable_name
        PRINT @cmd
        exec(@cmd)

    FETCH NEXT FROM fkeys_cursor INTO @fktable_name, @fk_name, 
    @fkcolumn1_name, 
    @fkcolumn2_name, 
    @fkcolumn3_name, 
    @fkcolumn4_name, 
    @fkcolumn5_name, 
    @fkcolumn6_name, 
    @fkcolumn7_name, 
    @fkcolumn8_name, 
    @fkcolumn9_name, 
    @fkcolumn10_name, 
    @fkcolumn11_name, 
    @fkcolumn12_name, 
    @fkcolumn13_name, 
    @fkcolumn14_name, 
    @fkcolumn15_name, 
    @fkcolumn16_name, 
    @pktable_name


END

CLOSE fkeys_cursor
DEALLOCATE fkeys_cursor


GO

/******END**********/