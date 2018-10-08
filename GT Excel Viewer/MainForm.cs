using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace GT_Excel_Viewer
{

    public partial class MainForm : DevExpress.XtraEditors.XtraForm
    {
        public MainForm()
        {
            InitializeComponent();
        }

        List<ExcelInfo> EI = new List<ExcelInfo>();
        string FileName = string.Empty;

        private void btnLoadData_Click(object sender, EventArgs e)
        {
            try
            {
                FileName = string.Empty;
                EI.Clear();
                
                ResetAllGrids();

                GetFilePath();

                if (!String.IsNullOrEmpty(FileName))
                {
                    LoadExcelFile();    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GetFilePath()
        {
            OpenFileDialog fDialog = new OpenFileDialog()
            {
                Filter = "",
                Multiselect = false,
            };

            fDialog.ShowDialog();

            FileName = fDialog.FileName;
        }

        private void ResetResultGrid()
        {
            GridControl1.BeginUpdate();
            GridControl1.DataSource = null;
            GridView1.Columns.Clear();
            GridControl1.EndUpdate();
        }

        private void ResetAllGrids()
        {
            ResetResultGrid();

            GridControl2.BeginUpdate();
            GridControl2.DataSource = null;
            GridView2.Columns.Clear();
            GridControl2.EndUpdate();
        }

        private void LoadExcelFile()
        {
            if (String.IsNullOrEmpty(FileName)) throw new Exception("Select file");

            OleDbConnection connExcel = new OleDbConnection();
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter oda = new OleDbDataAdapter();

            connExcel = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + FileName + ";Extended Properties=Excel 12.0;");
            cmdExcel.Connection = connExcel;
            connExcel.Open();
            System.Data.DataTable dt = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            if (dt == null) throw new Exception("File read error");

            foreach (DataRow row in dt.Rows)
            {
                string t = row["TABLE_NAME"].ToString();
                EI.Add(new ExcelInfo { ID = t, Sheet = CleanUpSheetName(t) });
            }

            CleanUpSheetName(EI[0].ID);

            connExcel.Close();
            
            GridControl1.BeginUpdate();
            GridControl1.DataSource = null;
            GridView1.Columns.Clear();

            GridControl1.DataSource = Ext.ToDataTable(EI);

            GridView1.Columns["ID"].Visible = false;

            GridView1.ClearSelection();
            GridControl1.EndUpdate();
        }

        string CleanUpSheetName(string sName)
        {
            int StartCharPos = sName.IndexOf("_");
            int EndCharPos = sName.IndexOf("$");
            return sName.Substring(StartCharPos + 1, EndCharPos - (StartCharPos + 2)).ToString().Trim();
        }

        private void GridView1_DoubleClick(object sender, EventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView view = (DevExpress.XtraGrid.Views.Grid.GridView)sender;
            Point pt = view.GridControl.PointToClient(Control.MousePosition);
            DoRowDoubleClick(view, pt);
        }

        private void DoRowDoubleClick(DevExpress.XtraGrid.Views.Grid.GridView view, Point pt)
        {
            try
            {
                DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo info = view.CalcHitInfo(pt);
                if (info.InRow || info.InRowCell)
                {
                    if (info.Column == null)
                    {
                    }
                    else
                    {
                        if (view.GetRowCellValue(info.RowHandle, "ID") == DBNull.Value) return;

                        string sheet = view.GetRowCellValue(info.RowHandle, "ID").ToString();

                        GetDetils(sheet);

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GetDetils(string sheet)
        {
            if (String.IsNullOrEmpty(FileName)) throw new Exception("Select file");

            GridControl2.BeginUpdate();
            GridControl2.DataSource = null;
            GridView2.Columns.Clear();

            OleDbConnection connExcel = new OleDbConnection();
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter oda = new OleDbDataAdapter();
            System.Data.DataSet ds = new System.Data.DataSet();

            connExcel = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + FileName + ";Extended Properties=Excel 12.0;");

            oda = new System.Data.OleDb.OleDbDataAdapter("select * FROM [" + sheet + "]", connExcel);
            oda.TableMappings.Add("Table", "Excel Data");
            oda.Fill(ds);

            connExcel.Close();

            GridControl2.DataSource = ds.Tables[0];

            GridView2.ClearSelection();
            GridControl2.EndUpdate();
        }
    
    }

    public class ExcelInfo
    {
        public string ID { get; set; }
        public string Sheet { get; set; }
    }

    public static class Ext
    {

        internal static DataTable ToDataTable<T>(this IList<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable dt = new DataTable();
            for (int i = 0; i <= properties.Count - 1; i++)
            {
                PropertyDescriptor property = properties[i];
                //dt.Columns.Add(property.Name, property.PropertyType);
                dt.Columns.Add(property.Name, Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType);
            }
            object[] values = new object[properties.Count - 1 + 1];
            foreach (T item in data)
            {
                for (int i = 0; i <= values.Length - 1; i++)
                    values[i] = properties[i].GetValue(item);
                dt.Rows.Add(values);
            }
            return dt;
        }

    }

}