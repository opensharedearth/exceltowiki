using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace exceltowiki
{
    class Program
    {
        static string s_inputFile = null;
        static string s_outputFile = null;
        static string[] s_columns = null;
        static string s_worksheetName = null;
        static int s_worksheetIndex = 1;
        static string[] s_formats = null;
        static bool s_overwrite = false;
        static bool s_headers = false;
        static string s_dateFormat = "g";

        static void Main(string[] args)
        {
            try
            {
                for (int i = 0; i < args.Length; ++i)
                {
                    string arg = args[i];
                    switch (arg)
                    {
                        case "-columns":
                            s_columns = GetColumnsArg(args, ++i);
                            break;
                        case "-formats":
                            s_formats = GetFormatsArg(args, ++i);
                            break;
                        case "-Y":
                            s_overwrite = true;
                            break;
                        case "-headers":
                            s_headers = true;
                            break;
                        case "-dateFormat":
                            s_dateFormat = GetDateFormatArg(args, ++i);
                            break;
                        case "-worksheet":
                            s_worksheetName = GetWorksheetName(args, ++i);
                            break;
                        default:
                            if (String.IsNullOrEmpty(arg)) throw new ApplicationException("Null argument seen on command line.");
                            if (arg[0] == '-') throw new ApplicationException($"Unrecognized switch argument {arg} seen.");
                            if (s_inputFile == null) s_inputFile = arg;
                            else if (s_outputFile == null) s_outputFile = arg;
                            else throw new ApplicationException($"Extraneous argument {arg} seen.");
                            break;
                    }
                }
                if (s_inputFile == null) throw new ApplicationException("Missing input file argument.");
                if (s_outputFile == null) throw new ApplicationException("Missing output file argument.");
                if (!File.Exists(s_inputFile)) throw new ApplicationException($"Input file '{s_inputFile}' does not exist.");
                if (File.Exists(s_outputFile) && !s_overwrite) throw new ApplicationException($"Output file '{s_outputFile} exists and overwrite not specified.");
                ConvertExcelToWiki(s_inputFile, s_outputFile, s_columns, s_formats, s_headers, s_dateFormat, s_worksheetName);
            }
            catch(Exception ex)
            {
                Usage(ex.Message);
            }
        }

        private static string GetWorksheetName(string[] args, int iarg)
        {
            string arg = iarg < args.Length ? args[iarg] : "";
            if (string.IsNullOrEmpty(arg) || arg[0] == '-') throw new ApplicationException("Missing worksheet name argument");
            return arg;
        }

        private static string GetDateFormatArg(string[] args, int iarg)
        {
            string arg = iarg < args.Length ? args[iarg] : "";
            if (string.IsNullOrEmpty(arg) || arg[0] == '-') throw new ApplicationException("Missing date format argument");
            try
            {
                string s = DateTime.Now.ToString(arg);
            }
            catch(Exception ex)
            {
                throw new ApplicationException("Invalid date format argument", ex);
            }
            return arg;
        }

        private static void ConvertExcelToWiki(string s_inputFile, string s_outputFile, string[] columns, string[] formats, bool headers, string dateFormat, string worksheetName)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                using (FileStream outputStream = File.Open(s_outputFile, FileMode.OpenOrCreate | FileMode.Truncate, FileAccess.Write, FileShare.None))
                {
                    using(StreamWriter writer = new StreamWriter(outputStream))
                    {
                        string path = s_inputFile;
                        if (!Path.IsPathRooted(path)) path = Path.Combine(Directory.GetCurrentDirectory(), s_inputFile);
                        Workbook workbook = app.Workbooks.Open(path);
                        if (workbook != null)
                        {
                            if (worksheetName == null) worksheetName = "1";
                            Worksheet worksheet = workbook.Worksheets[worksheetName];
                            if (worksheet != null)
                            {
                                writer.WriteLine("{| class=\"wikitable\"");
                                int nrows = worksheet.UsedRange.Rows.Count;
                                if (nrows == 0) throw new ApplicationException("The source table is empty.");
                                int ncolumns = worksheet.UsedRange.Columns.Count;
                                if (columns == null) columns = GetExcelColumns(ncolumns);
                                if (ncolumns < columns.Length) throw new ApplicationException("The source table does not have enough columns");
                                if (formats == null) formats = new string[0];
                                int irow = 1;
                                if(s_headers)
                                {
                                    writer.WriteLine("|+");
                                    foreach(string column in columns)
                                    {
                                        string s = worksheet.Cells[irow, column].Value;
                                        writer.WriteLine("|" + s);
                                    }
                                    irow++;
                                }
                                for (int r = irow; r <= nrows; ++r)
                                {
                                    writer.WriteLine("|-");
                                    for (int i = 0; i < columns.Length; ++i)
                                    {
                                        string column = columns[i];
                                        string format = i < formats.Length ? formats[i] : "";
                                        string s = worksheet.Cells[r, column].Value;
                                        string s1 = FormatColumn(s, format, dateFormat);
                                        writer.WriteLine("|" + s1);
                                    }
                                }
                                writer.WriteLine("|}");
                            }
                            else
                            {
                                throw new ApplicationException($"Worksheet {worksheetName} not found.");
                            }
                        }
                        writer.Close();
                    }
                }
            }
            finally
            {
                app.Quit();
            }
        }

        private static string[] GetExcelColumns(int ncolumns)
        {
            if (ncolumns > 26 * 26) throw new ApplicationException("Source table has too many columns.");
            List<string> columns = new List<string>();
            for(int i = 0; i < ncolumns; i++)
            {
                char a = (char)((int)'A' + i % 26);
                if(i < 26)
                {
                    columns.Add(new string(a, 1));   
                }
                else
                {
                    char b = (char)((int)'A' + i / 26);
                    columns.Add(new string(new char[] { b, a }));
                }
            }
            return columns.ToArray();
        }

        private static string FormatColumn(string s, string format, string dateFormat)
        {
            switch (format)
            {
                case "DATE": return FormatDate(s, dateFormat);
                case "HTML": return FormatHTML(s);
                default:
                    return s;
            }
        }

        private static string FormatHTML(string s)
        {
            string[] p = GetParagraphs(s);
            StringBuilder sb = new StringBuilder();
            foreach(string s0 in p)
            {
                sb.AppendLine(s0);
                sb.AppendLine();
            }
            return sb.ToString();
        }

        private static string[] GetParagraphs(string s)
        {
            List<string> p = new List<string>();
            int i = 0;
            while(i < s.Length)
            {
                int j = s.IndexOf("<p>", i);
                if (j < 0)
                {
                    p.Add(s.Substring(i).Trim());
                    i = s.Length;
                }
                else if(j == i)
                {
                    int k = s.IndexOf("</p>", i);
                    if (k < 0)
                    {
                        p.Add(s.Substring(i + 3).Trim());
                        i = s.Length;
                    }
                    else
                    {
                        p.Add(s.Substring(i + 3, k - i - 3).Trim());
                        i = k + 4;
                    }
                }
                else
                {
                    i = j;
                }
            }
            return p.ToArray();
        }

        private static string FormatDate(string s, string dateFormat)
        {
            if(DateTime.TryParse(s, out DateTime v))
            {
                return v.ToString(dateFormat);
            }
            else
            {
                return s;
            }
        }

        private static string[] GetColumnsArg(string[] args, int iarg)
        {
            string arg = iarg < args.Length ? args[iarg] : "";
            if (string.IsNullOrEmpty(arg) || arg[0] == '-') throw new ApplicationException("Missing columns argument");
            string[] columns = arg.Split(',');
            foreach(var column in columns)
            {
                if (!IsValidExcelColumn(column)) throw new ApplicationException("Invalid Excel column seen; must be A-Z and AA-ZZ");
            }
            return columns;
        }

        private static bool IsValidExcelColumn(string column)
        {
            string column1 = column.ToUpper();
            string validColumns = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            if (column1.Length < 1 || column1.Length > 2) return false;
            if (column1.Length == 1 && validColumns.Contains(column1.ToUpper())) return true;
            else if (validColumns.Contains(column1[0]) && validColumns.Contains(column[1])) return true;
            else return false;
        }

        private static string[] GetFormatsArg(string[] args, int iarg)
        {
            string arg = iarg < args.Length ? args[iarg] : "";
            if (string.IsNullOrEmpty(arg) || arg[0] == '-') throw new ApplicationException("Missing formats argument");
            string[] formats = arg.ToUpper().Split(',');
            foreach(var format in formats)
            {
                switch (format)
                {
                    case "":
                    case "HTML":
                    case "DATE":
                        break;
                    default:
                        throw new ApplicationException("Invalid formats argument");
                }
            }
            return formats;
        }

        static void Usage(string error = "")
        {
            if(!String.IsNullOrEmpty(error))
            {
                Console.WriteLine("Error: " + error);
                Console.WriteLine();
            }
            Console.WriteLine("Usage:");
            Console.WriteLine();
            Console.WriteLine("exceltowiki [-Y] input [-columns clist] [-format flist] [-sheet sname] [-dateFormat dformat] [-headers] output");
            Console.WriteLine();
            Console.WriteLine("\tinput       Path to input Excel spreadsheet file (.xls)");
            Console.WriteLine("\toutput      Path to output wiki contents file (.wiki)");
            Console.WriteLine("\t-Y          Overwrite destination file if it exists");
            Console.WriteLine("\t-columns    Comma separated list of column names, e.g. A,C,D,AA,B");
            Console.WriteLine("\t-format     Comma separated list of format conversions, e.g. date,,html");
            Console.WriteLine("\t-headers    Columns have headers");
            Console.WriteLine("\t-dateFormat Format of date output on date conversion of column data");
            Console.WriteLine();
            Console.WriteLine("Where:");
            Console.WriteLine();
            Console.WriteLine("\tclist       Excel column list, 1 or 2 alphabetic letters in order that the columns will appear in output.");
            Console.WriteLine("\tflist       Format list, one of date or html or empty string. Position in list corresponds to position in clist.");
            Console.WriteLine("\t            date format indicates column is date/time.  Converted in output according to date-format argument");
            Console.WriteLine("\t            html format indicates column contains HTML  Converted to wikitext format.");
            Console.WriteLine("\t                    only the <p> tag is currently supported");
            Console.WriteLine("\tdformat     Standard or custom datetime string as described in .NET Documentation. e.g. dd-MMM-yy");

        }
    }
}
