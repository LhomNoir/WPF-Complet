/*
namespace WpfApp1
{
    class DataTable
    {
        private static SQLDataBase sqlDB = new SQLDataBase("fr-sql.hilti.com", "Junk");
        //private static System.Data.DataTable dataTableArg = null;                            
        private static string schema = "[voi]";
        private static string tableNameArg = "[AAA_Destock_Tool]";


        static void Main(string[] args)
        {

            // --  --
            CreateVioTable();

            Console.WriteLine("Press ENTER to exit...");
            Console.ReadLine();
        }


        private static void CreateVioTable()
        {
            var dataTableArg = null;
            dataTableArg = GetOptionDataTable(@"C:\Users\mabomic\Desktop\VOI inputs.xlsx");

            //string assemblyName = System.Reflection.Assembly.GetEntryAssembly()?.GetName()?.Name ?? ".Net Application";
            //string connectionString = "Data Source=fr-sql.hilti.com;Initial Catalog=Junk;Application Name=assemblyName;Connection Timeout=90;Integrated Security=True";

            using (SQLTable voiTable = new SQLTable(sqlDB, "voi", "AAA_Destock_Tool"))
            //using (SqlConnection connection = new SqlConnection(connectionString))
            {
                voiTable.Import(dataTableArg);

                // --  Truncate table  --
                voiTable.Truncate();

                //voiTable.WriteToServer(dataTableArg);
                voiTable.WriteToServerTL();

                #region MyRegion
                //// Create a table with some rows. 
                //System.Data.DataTable voiData = dataTableArg;

                //// Create the SqlBulkCopy object. 
                //// Note that the column positions in the source DataTable 
                //// match the column positions in the destination table so 
                //// there is no need to map columns. 
                //using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(connection))
                //{
                //    sqlBulkCopy.BulkCopyTimeout = 15 * 60;  //-> 15 min !
                //    sqlBulkCopy.DestinationTableName = string.Format("{0}.{1}", schema, tableNameArg);
                //    sqlBulkCopy.BatchSize = dataTableArg.Rows.Count;

                //    try
                //    {
                //        connection.Open();

                //        // Write from the source to the destination.
                //        sqlBulkCopy.WriteToServer(voiData);
                //    }
                //    catch (Exception ex)
                //    {
                //        Console.WriteLine(ex.Message);
                //    }
                //} 
                #endregion
            }
        }

        #region --  Create data table  --
        private static System.Data.DataTable GetOptionDataTable(string filename)
        {
            Application xlApp;
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;

            var missing = System.Reflection.Missing.Value;

            xlApp = new Application();
            xlWorkBook = xlApp.Workbooks.Open(filename, false, true, missing, missing, missing, true, XlPlatform.xlWindows, '\t', false, false, 0, false, true, 0);
            //xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(2);

            Range xlRange = xlWorkSheet.UsedRange;
            Array myValues = (Array)xlRange.Cells.Value2;

            #region --  Create datatable  --  
            int rowCount = myValues.GetLength(0);
            int columnCount = myValues.GetLength(1);

            System.Data.DataTable myDataTable = new System.Data.DataTable();
            myDataTable.Clear();

            // --- Get header information  ---
            for (int i = 1; i <= columnCount; i++)
            {
                myDataTable.Columns.Add(new System.Data.DataColumn(myValues.GetValue(1, i).ToString()));
            }

            // --- Get the row information ---
            for (int a = 2; a <= rowCount; a++)
            {
                object[] poop = new object[columnCount];
                for (int b = 1; b <= columnCount; b++)
                {
                    poop[b - 1] = myValues.GetValue(a, b);
                }
                System.Data.DataRow row = myDataTable.NewRow();
                row.ItemArray = poop;
                myDataTable.Rows.Add(row);
            }

            //// --- Get header information  ---
            //for (int i = 1; i <= columnCount - 1; i++)
            //{
            //    myDataTable.Columns.Add(new DataColumn(myValues.GetValue(1, i).ToString()));
            //}

            //// --- Get the row information ---
            //for (int a = 2; a <= rowCount - 1; a++)
            //{
            //    object[] poop = new object[columnCount - 1];
            //    for (int b = 1; b <= columnCount - 1; b++)
            //    {
            //        poop[b - 1] = myValues.GetValue(a, b);
            //    }
            //    DataRow row = myDataTable.NewRow();
            //    row.ItemArray = poop;
            //    myDataTable.Rows.Add(row);
            //}
            #endregion

            xlWorkBook.Close(true, missing, missing);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            return myDataTable;
        }

        // -- Release object --
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        #endregion 
    }
    
}
*/

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace ReadExcel
{
    class Program00
    {
        public static void DateTime00()
        {
            string file = @"D:\tmp\Store 29-09-15.xlsx";

            var dataSet = GetDataSetFromExcelFile(file);

            Console.WriteLine(string.Format("reading file: {0}", file));
            Console.WriteLine(string.Format("coloums: {0}", dataSet.Tables[0].Columns.Count));
            Console.WriteLine(string.Format("rows: {0}", dataSet.Tables[0].Rows.Count));
            Console.ReadKey();
        }

        private static string GetConnectionString(string file)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            string extension = file.Split('.').Last();

            if (extension == "xls")
            {
                //Excel 2003 and Older
                props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
                props["Extended Properties"] = "Excel 8.0";
            }
            else if (extension == "xlsx")
            {
                //Excel 2007, 2010, 2012, 2013
                props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
                props["Extended Properties"] = "Excel 12.0 XML";
            }
            else
                throw new Exception(string.Format("error file: {0}", file));

            props["Data Source"] = file;

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        private static DataSet GetDataSetFromExcelFile(string file)
        {
            DataSet ds = new DataSet();

            string connectionString = GetConnectionString(file);

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }

            return ds;
        }
    }
}
