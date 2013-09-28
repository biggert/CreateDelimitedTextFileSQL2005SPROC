Create Delimited Text File (SQL 2005 Stored procedure)
Introduction

Over the past 2 years, I’ve primarily worked with importing/exporting data into multiple formats. One product I worked on in particular could import and export data into over 20 different formats while not only staying within a reasonable time limit and memory usage but also giving the user multiple options for specifying the source/destination mappings as well as other elements used in the application. The focus of this application was speed and low memory usage and in the end, it could import a 3GB CSV file to a SQL database in less than an hour while staying below 60 MB of ram usage on the machine.

After moving on, I’ve continued to explore the knowledge of bulk loading data and decided to write a quick and easy script that will handle this on a SQL server in the quickest way possible not only as a testament to myself but also as a useful tool that can be easily implemented by other SQL administrators and programmers when needed. Given how much I’ve learned from the open source community, I wanted to give back with this little stored procedure. In addition to this, I have at least 2 more projects in mind that I hope to release soon that will definitely fill some gaps in the data industry.

If you have any problems, questions, or just comments about this stored procedure, feel free to shoot me an email. Also, if you have any changes/addition you’d like to add, I’ll be happy to take your suggestions and add it in.

To contact me, click here

For a quick download link, click here.

Description

This stored procedure allows the user to create a delimited text file from a SQL table, SQL view, Microsoft Access Database File, Delimited Text File, Microsoft Excel Spreadsheet File, or a DBase III (DBF) File. It is a fully self-contained stored procedure that needs to no other functions and/or stored procedures to run. It is run directly on the SQL server (only 2005 for now) and makes connections to the source/destination files from the server. It allows SQL-style criteria to be used against any of the supported data sources. It also allows the user to specify the format of the delimited text file that it creates. In addition to this, the user can also specify which columns should be exported to the delimited text files. By default, all columns will be exported.

Supported Source Data Types

SQL Table
SQL View
Microsoft Access Database File
Microsoft Excel Spreadsheet File
DBASE III File
Delimited Text File
Requirements

SQL Server 2005
Must have the xp_cmdshell functionality enabled on the server. This is used to execute the bcp utility which provides the fastest way of exporting data to the delimited text file destination.
If exporting data from a source type other than SQL, OPENROWSET functionality must be enabled on the server.
Parameters

Here is a list of available parameters:

Source (required)
Specifies the source table, view, or file
When specifying a file, always use the full file path and file name
DestinationFile (required)
Specifies the destination delimited text file
ColumnList (defaults to all)
Allows a comma delimited string to be provided which will cause it only to export the specified columns in the table. If left blank, it will get all columns. If you do specify a column list, make sure the columns exist in the source or it will error out
Delimiter (defaults to a comma [,] )
Specifics the delimiter for the delimited text file (Use “TAB” to specify a TAB character)
Qualifier (defaults to a double-quote character ["] )
Specifies the text qualifier for the delimited text file
Criteria (defaults to blank)
Specifies the criteria to be used against the source. Do not include the WHERE keyword if using this argument
FirstRow (defaults to 0)
Specifies the first row to begin in the source table. Anything greater than 0 will cause the delimited text file to be written without a header row
LastRow (defaults to 0)
Specifies the last row to export in the source table. Zero here means get them all
Username (defaults to blank)
Specifies the username for logging into the SQL server. Leave blank to use a trusted connection
Password (defaults to blank)
Specifies the password for logging into the SQL server.
Server (defaults to blank)
Specifies the server to connect to. Leave blank to use the default instance on the SQL server from which the stored procedure is run
Source Type (defaults to SQL)
Used to specify the source type
If left blank, this stored procedure will attempt to figure out the source type by searching the extension of the filename (MDB – Access, XLS – Excel, DBF – DBase III, CSV – Delimited, etc.)
Legal values are as follows:
SQL
SQL
Delimited Text File
Delimited
Text
CSV
TXT
Dbase III
DBF
DBASE
DBASE3
DBASEIII
DBASE 3
DBASE III
FOXPRO
Access
Access
MDB
Excel
XLS
Excel
SourceTableName (defaults to Source)
Only used when exporting data from a Microsoft Access Database File or Microsoft Excel Spreadsheet File. More specifically, this is used to identify the source table or spreadsheet from which to export data
OtherConnection (defaults to blank)
This is not used and should always be left blank. I kept this in as to add functionality later on. I’ve got a horrible memory so if I took it out, I wouldn’t remember to work on it later
Benchmarks

I ran most of my benchmarks against a database without any indexes and only 1 million rows. Anything less than 25 fields and 200,000 rows took less than 5 seconds (the stored procedure took longer to run than the BCP process )

500,000 rows
10 fields
Time: 23 seconds
File Size: 51MB
21932 rows per second
1 million rows
10 fields
Time: 42 seconds
File Size: 96MB
24178 rows per second
1 million rows
32 fields
Time: 2 minutes, 24 seconds
File Size: 258 MB
6927 rows per second
1 million rows
110 fields
Time: 13 minutes
File Size: 826 MB
1288 rows per second
Technical Details

Basically, this stored procedure creates a temporary view and runs the bcp command line tool against the view. After it is finished (or an error occurs), it will delete the temporary view. Also, the temporary view has a unique name so it will not conflict with itself if the stored procedure is run multiple times simultaneously. If an error does occur, it should tell you why. The bcp command line tool is provided by Microsoft and comes with SQL server. As long as the SQL server has access to run the bcp tool (which it should by default) and has access to create a file in the destination folder, it will run fine.

Example

Create a delimited text file from a SQL table with a comma (,) as the delimiter and double quotes (”) as the text qualifier.
exec sp_CreateDelimitedTextFile ‘TableName’, ‘C:\destination.csv’
Create a delimited text file from a SQL table with a pipe (|) as the delimiter and no text qualifiers
exec sp_CreateDelimitedTextFile ‘TableName’, ‘C:\destination.txt’, @Delimiter=’|’, @Qualifier=”
Create a delimited text file from a SQL table using criteria
exec sp_CreateDelimitedTextFile ‘TableName’, ‘C:\destination.csv’, @Criteria=’ Salary > 40000 ‘
Create a delimited text file from a SQL table exporting only 3 columns
exec sp_CreateDelimitedTextFile ‘TableName’, ‘C:\destination.csv’, @ColumnList=’Column1,Column2,Column3′
Create a delimited text file from a SQL table only exporting the first 100 results and omitting the header row
exec sp_CreateDelimitedTextFile ‘TableName’, ‘C:\destination.csv’, @FirstRow=1, @LastRow=101
Create a delimited text file from a delimited text file with a pipe (|) as the delimiter and no text qualifiers
exec sp_CreateDelimitedTextFile ‘C:\source.csv’, ‘C:\destination.txt’, @Delimiter=’|’, @Qualifier=”
exec sp_CreateDelimitedTextFile ‘C:\source.dat’, ‘C:\destination.txt’, @Delimiter=’|’,@Qualifier=”, @SourceType=’Delimited’
Create a delimited text file from an access database table
exec sp_CreateDelimitedTextFile ‘C:\access.mdb’,’C:\destination.csv’,@SourceTableName=’AccessTable’
Create a delimited text file from a DBase III file
exec sp_CreateDelimitedTextFile ‘C:\dbase.dbf’, ‘C:\destination.csv’
exec sp_CreateDelimitedTextFile ‘C:\dbase.dat’, ‘C:\destination.csv’, @SourceType=’Dbase’
Create a delimited text file from an Excel spreadsheet file
exec sp_CreateDelimitedTextFile ‘C:\excel.xls’,’C:\destination.csv’, @SourceTableName=’Sheet1′
Current Issues

The most important bug found so far is that, because numeric field are being converted to varchar, there seems to be a loss of precision for large floats. I believe it lops off the decimal value after the 6th or 7th place AFTER the decimal. Also, it will only read from comma-delimited text files.

I haven’t run this stored procedure through the ringer and made sure every data type works. Other than that, most things can be contributed to user error and, although it’d look a lot prettier if I built a few more try catches to fix invalid parameter values, it’s not like I’m getting paid to write this thing. If time comes along, I might make it a little prettier. I know it can be a bit more optimized (like unnecessarily using varchar(max)) but I was looking for 100% functionality versus memory/speed. Plus, it’s pretty fast as is.

I had a visitor report a problem and when fixing this, I found 2 additional problems that I will work on in the near future. First, it doesn’t seem to like the ntext data type. I haven’t tested if this is true for the text and image data type but since they are the same in the underlying SQL, they may have problems. Secondly, specifying the database name before the table will not work. It is coded in a way where the stored procedure assumes the database which is being used to run the stored procedure contains the table you are wanted to export. I will update this functionality in the future to allow a database name to be specified through the normal way such as “[ExampleDB].dbo.[ExampleTable]” and/or with a parameter. In the meantime, just run the “USE [Example Database]” command before executing the stored procedure if possible. (Feel free to fix it yourself though by doing a find replace on db_name() with your own database name)

Download Script

Update (10/22/2007): For those having problems with the permissions needed on the SQL server to execute this stored procedure, run the following script:

EXEC sp_configure 'show advanced options', 1
GO
RECONFIGURE
GO
EXEC sp_configure 'xp_cmdshell', 1
GO
RECONFIGURE
GO
