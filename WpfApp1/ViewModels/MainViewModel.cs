using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.IO;
using System.Linq;
using WpfApp1.Commands;
using WpfApp1.Models;
using WpfApp1.ViewModels.Helpers;

namespace WpfApp1.ViewModels
{
    public class MainViewModel : NotifyUIBase
    {
        #region --  Fields --
        private string Region_FilePath = @"C:\Users\MMabou\Desktop\TEST Kantar\region.csv";
        #endregion

        #region --  Properties --

        public ObservableCollection<Region> Regions { get; set; }
        private Dictionary<string, string> RegionDictionary = new Dictionary<string, string>();

        private Region selectedRegion;
        public Region SelectedRegion
        {
            get
            {
                return selectedRegion;
            }
            set
            {
                selectedRegion = value;
                RaisePropertyChanged();
            }
        }

        #endregion

        #region --  Constructor --
        public MainViewModel()
        {
            ImportExcelFileToDataTable();

            // *********
            Regions = new ObservableCollection<Region>();

            GetRegions();

            GetRegionPopulationCommand = new DelegateCommand(x => GetRegionPopulation());
        }
        #endregion

        #region -- Commands --
        public DelegateCommand GetRegionPopulationCommand { get; set; }
        #endregion

        #region --  Methodes --
        private void GetRegions()
        {
            Regions.Clear();
            RegionDictionary.Clear();
            var regions = File.ReadAllLines(Region_FilePath).Select(a => a.Split(';')).ToList();
            foreach (var region in regions)
            {
                RegionDictionary.Add(region[0], region[1]);
                Regions.Add(new Region { RegionName = region[0], RegionPopulation = region[1] });
            }

            using (var reader = new StreamReader(Region_FilePath))
            {
                Regions.Clear();
                while (!reader.EndOfStream)
                {
                    var values = reader.ReadLine().Split(';');
                    Regions.Add(new Region { RegionName = values[0], RegionPopulation = values[1] });
                }
            }
        }

        public void GetRegionPopulation()
        {
            if (SelectedRegion != null)
            {
                var selectedRegion = SelectedRegion.RegionName;
                var population = RegionDictionary.ContainsKey(selectedRegion);
            }

            //Console.WriteLine(population);
        }
        #endregion

        #region ------------------------------------------
        private void ImportExcelFileToDataTable()
        {
            try
            {
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='c:\\csharp.net-informations.xls';Extended Properties=Excel 8.0;");
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection);
                MyCommand.TableMappings.Add("Table", "TestTable");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                //dataGridView1.DataSource = DtSet.Tables[0];
                var datas = DtSet.Tables[0];
                MyConnection.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private System.Data.DataTable ReadExcelFile(string sheetName, string path)
        {
            using (System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection())
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                string Import_FileName = path;
                string fileExtension = Path.GetExtension(Import_FileName);
                if (fileExtension == ".xls")
                    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";
                if (fileExtension == ".xlsx")
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                using (System.Data.OleDb.OleDbCommand comm = new System.Data.OleDb.OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";

                    comm.Connection = conn;

                    using (System.Data.OleDb.OleDbDataAdapter da = new System.Data.OleDb.OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(null);
                        //da.Fill(dt, "");
                        return dt;
                    }

                }
            }
        }

        private void Import_To_Grid(string FilePath, string Extension, string isHDR)
        {
            string conStr = "";

            switch (Extension)
            {
                case ".xls": //Excel 97-03  
                    conStr = ConfigurationManager.ConnectionStrings["Excel03ConString"]
                    .ConnectionString;
                    break;
                case ".xlsx": //Excel 07  
                    conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"]
                    .ConnectionString;
                    break;
            }

            conStr = String.Format(conStr, FilePath, isHDR);
            System.Data.OleDb.OleDbConnection connExcel = new System.Data.OleDb.OleDbConnection(conStr);
            System.Data.OleDb.OleDbCommand cmdExcel = new System.Data.OleDb.OleDbCommand();
            System.Data.OleDb.OleDbDataAdapter oda = new System.Data.OleDb.OleDbDataAdapter();
            System.Data.DataTable dt = new System.Data.DataTable();
            cmdExcel.Connection = connExcel;

            ////Get the name of First Sheet  
            //connExcel.Open();
            //DataTable dtExcelSchema;
            //dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            //string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
            //connExcel.Close();

            ////Read Data from First Sheet  
            //connExcel.Open();
            //cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
            //oda.SelectCommand = cmdExcel;
            //oda.Fill(dt);
            //connExcel.Close();

            //Bind Data to GridView  
            //GridView1.Caption = Path.GetFileName(FilePath);
            //GridView1.DataSource = dt;
            //GridView1.DataBind();
        }

        #endregion
    }
}

