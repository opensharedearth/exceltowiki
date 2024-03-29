# exceltowiki
Excel to Wikitext converter

Usage:

exceltowiki input [--columns clist] [--format flist] [--worksheet sname] [--date-format dformat] [--headers]
                  [--wiki wikiurl --username username --password password] [--overwrite]
                  [--table-title title | --no-table] [--title-column tcolumn]
                  [--page-prefix prefix] [--no-pages] [--page-columns pclist] [--page-format pflist]
                  [--edit-summary summary]

	input         Path to input Excel spreadsheet file (.xls)
	--columns     Comma separated list of column names, e.g. A,C,D,AA,B
	--format      Comma separated list of format conversions, e.g. date,,date
	--worksheet   The name or index of the worksheet in the Excel workbook to be used as a source
	--worksheet-range   The range of cells to use for the table.  By default the used range of the worksheet
	--headers     Columns have headers
	--date-format Format of date output on date conversion of column data
	--wiki        Destination wiki
	--username    Username for wiki login
	--password    Password for wiki login (bot password)
	--overwrite   Overwrite wiki pages if they already exist
	--table-title The title of the wiki page containing the table
	--no-table    Do not write the table into the wiki
	--title-column    The column in the Excel spreadsheet that contains the title of the wiki page to be created
	--no-pages    Pages are not created.  Links are still inserted in table if title-column defined.
	--page-prefix The prefix that is prepended to the title of each wiki page created except for the table page
	--page-columns    Comma separated list of Excel columns that become wiki page sections. Section name taken from column header.
	--page-format Comma separated list of formats for excel columns that become wiki page sections
	--edit-summary    The description of the page edit.  By default "Modified by exceltowiki"
	--help        display usage
	--version     display version

Where:

	clist         Comma separated Excel column list, 1 or 2 alphabetic letters in order that the columns will appear in output.
	flist         Comma separated format list, one of date or empty string. Position in list corresponds to position in clist.
	              date format indicates column is date/time.  Converted in output according to date-format argument
	sname         Excel worksheet name or index.  Sheet indicies are 1-based.
	dformat       Standard or custom datetime string as described in .NET Documentation. e.g. "dd-MMM-yy".  Default is "g"
	wikiurl       The absolute URL of the wiki destination for the pages created
	username      The username for the wiki indicated by wikiurl
	password      The password for the wiki indicated by wikiurl; this must be the special bot password
	title         The title of the wiki page that contains the table
	prefix        The prefix for all of the wiki pages created except for the table page
	pclist        The comma separated list of spreadsheet column that become sections in the wiki page
	pflist        The comma separated format list of the spreadsheet column that will become sections of the wiki page
	summary       The page edit description

