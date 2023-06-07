using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using Telerik.WinControls.UI;
using Telerik.WinControls.Enumerations;
using Telerik.WinControls.Export;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using Telerik.WinControls.UI.Export;
using TestingReportSetWithTelerik.Interfaces;
using TestingReportSetWithTelerik.Model;

namespace TestingReportSetWithTelerik
{
    public partial class RadForm1 : RadForm, IviewEvent
    {

        //private string connectionString = "Data Source=GGINC;Initial Catalog=Northwind;Integrated Security=True";
        //DataTable printDataTable = new DataTable();
        
        //DateTime date1 = DateTime.Now;
        //DateTime date2 = DateTime.Now;

        private IviewEvent viewEvent;
        private string sqlconecctionString;
        private IRepositorio repositorio;
        
        DataTable datatable = new DataTable();
        //ModelDate model = new ModelDate();

        DateTime IviewEvent.Date1 { get => radDate1.Value; set => radDate1.Value = value; }
        DateTime IviewEvent.Date2 { get => radDate2.Value; set => radDate2.Value = value; }

        public RadForm1(string sqlconecctionString, IRepositorio repository)
        {
            InitializeComponent();
            this.sqlconecctionString = sqlconecctionString;
            this.repositorio = repository;
           
        }

        private async void RadForm1_Load(object sender, EventArgs e)
        {
            radWaiting.Visible = false;

            //radDate1.ValueChanged += (s, ee) =>
            //{
            //    model.Date11 = radDate1.Value;
            //};

            //radDate2.ValueChanged += (s, ee) =>
            //{
            //    model.Date22 = radDate2.Value;
            //};
            //model.Date11 = viewEvent.Date1;
            //model.Date22 = viewEvent.Date2;

            radWaiting.Visible = true;
            radWaiting.StartWaiting();

            datatable = await repositorio.ShowGridview(radDate1.Value, radDate2.Value);

            

            if (datatable.Columns.Count > 0)
            {
                for (int i = 0; i < datatable.Columns.Count; i++)
                {
                    string columnName = datatable.Columns[i].ColumnName;
                    try
                    {
                        ListViewDetailColumn column = new ListViewDetailColumn(columnName);
                        column.HeaderText = columnName;
                        radListViewLeft.Items.Add(columnName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("columna ya existe" + ex);
                    }

                }

            }
            radWaiting.StopWaiting();
            radWaiting.Visible = false;
        }
        //public (DataTable,SqlDataReader, SqlCommand) ShowDatagrid()
        //{
        //    string columnName="";
        //    SqlConnection connection = new SqlConnection(connectionString);

        //    connection.Open();
        //    SqlCommand command = new SqlCommand(StoreprocedureName, connection);
        //    command.CommandType = CommandType.StoredProcedure;
        //    ParameterDate(command);
        //    SqlDataReader reader = command.ExecuteReader();
            
        //    if (reader.FieldCount > 0)
        //    {
        //        for (int i = 0; i < reader.FieldCount; i++)
        //        {
        //            columnName = reader.GetName(i);

        //            if (!datatable.Columns.Contains(columnName))
        //            {
        //                datatable.Columns.Add(columnName);
        //            }                  
                    

        //        }
        //        while (reader.Read())
        //        {
        //            DataRow row = datatable.NewRow();

        //            for (int i = 0; i < reader.FieldCount; i++)
        //            {
        //                columnName = reader.GetName(i);
        //                object value = reader.GetValue(i);
        //                row[columnName] = value;
        //            }
        //            datatable.Rows.Add(row);
        //        }
        //    }
        //    return (datatable, reader, command);
        //}

       

        private async void FillAndupdateDataGridView()
        {
            radWaiting.Visible = true;
            radWaiting.StartWaiting();

            datatable = await repositorio.ShowGridview(radDate1.Value, radDate2.Value);
            
            radWaiting.StopWaiting();
            radWaiting.Visible = false;

            foreach (ListViewDataItem item in radListViewRigth.Items)
            {
                string columnName = item.Text;

                if (GridView1.Columns[columnName] == null)
                {
                    DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                    newColumn.Name = columnName;
                    newColumn.HeaderText = columnName;

                    GridView1.Columns.Add(columnName);
                }
                
                
                if (datatable.Columns.Contains(columnName))
                {
                    var colunm = GridView1.Columns.FirstOrDefault(x => x.Name == columnName);
                    if (colunm == null) return;

                    // Obtener el índice de la columna2
                    int columnIndex = colunm.Index;


                    
                    List<string> rowValues = new List<string>(); // Almacenar temporalmente los valores de la fila
                    foreach (DataRow row in datatable.Rows)
                    {
                        string value = row[columnName].ToString();
                        
                        rowValues.Add(value); // Agregar el valor a la lista

                        // Asegurarse de tener suficientes filas en el RadGridView
                        int requiredRowCount = Math.Max(GridView1.Rows.Count, rowValues.Count);
                        AddMissingRows(GridView1, requiredRowCount);

                        // Agregar los valores en la fila correspondiente del RadGridView
                        AddValuesToRadGridView(GridView1, columnIndex, rowValues);
                    }
                                   

                    

                }

                GridView1.Columns[columnName].Width = 150;                
            }
        }
        private async void radBTNrigth_Click(object sender, EventArgs e)
        {
           
            // Eliminar los elementos seleccionados en el ListView1 y agregarlos al ListView2
            List<ListViewDataItem> itemsToRemove = new List<ListViewDataItem>();
            // Mover los elementos seleccionados del ListView1 al ListView2
            foreach (ListViewDataItem checkedItem in radListViewLeft.CheckedItems)
            {
                string columnName = checkedItem.Text;
                radListViewRigth.Items.Add(columnName);
                itemsToRemove.Add(checkedItem);
            }
            // Eliminar los elementos de la colección original
            foreach (ListViewDataItem itemToRemove in itemsToRemove)
            {
                radListViewLeft.Items.Remove(itemToRemove);
            }


            GridView1.Columns.Clear();
            GridView1.Rows.Clear();

           

            FillAndupdateDataGridView();

            
            
        }


        private void AddMissingRows(RadGridView radGridView, int requiredRowCount)
        {
            while (radGridView.Rows.Count < requiredRowCount)
            {
                GridViewDataRowInfo newRow = (GridViewDataRowInfo)radGridView.Rows.AddNew();
                newRow.Cells[0].Value = string.Empty;
            }
        }
        private void AddValuesToRadGridView(RadGridView radGridView, int columnIndex, List<string> values)
        {
            for (int i = 0; i < values.Count; i++)
            {
                radGridView.Rows[i].Cells[columnIndex].Value = values[i];
            }
        }

        private static void ExportarDataTableAExcel(DataTable dataTable, string nombreArchivo)
        {
            if (dataTable == null || dataTable.Columns.Count == 0)
                throw new Exception("ExportToExcel: Null or empty input table!\n");

            var excelApp = new Excel.Application();
            excelApp.Workbooks.Add();

            Excel._Worksheet workSheet = (Excel._Worksheet)excelApp.ActiveSheet;
            // column headings
            for (var i = 0; i < dataTable.Columns.Count; i++)
            {
                workSheet.Cells[1, i + 1] = dataTable.Columns[i].ColumnName;
            }

            // rows
            for (var i = 0; i < dataTable.Rows.Count; i++)
            {
                // to do: format datetime values before printing
                for (var j = 0; j < dataTable.Columns.Count; j++)
                {
                    workSheet.Cells[i + 2, j + 1] = dataTable.Rows[i][j];
                }
            }
            workSheet.SaveAs(nombreArchivo);
            excelApp.Quit();
            MessageBox.Show("Excel file saved!");


            excelApp.Quit();

            MessageBox.Show("Se ha generado el documento");
        }


        private async void btnPrint_Click_1(object sender, EventArgs e)
        {
            //string nameFile = @"C:\Users\Dev_2\Desktop\archivo.xlsx";
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + $"\\archivo.xlsx";
            //ExportToExcel(nuevaTabla, nameFile);

            //var dataprint = ShowDatagrid();
            radWaiting.Visible = true;
            radWaiting.StartWaiting();

           Task.Run(()=> ExportarDataTableAExcel(this.datatable, path)).Wait();
           

            
            GridViewSpreadStreamExport spreadStreamExport = new GridViewSpreadStreamExport(this.GridView1);
            spreadStreamExport.ExportVisualSettings = true;
            spreadStreamExport.FileExportMode = FileExportMode.CreateOrOverrideFile;



            //SpreadStreamCellFormattingEventHandler SpreadStreamExport_CellFormatting = ;
            //spreadStreamExport.CellFormatting += new SpreadStreamCellFormattingEventHandler(SpreadStreamExport_CellFormatting);
            //spreadStreamExport.SummariesExportOption = SummariesOption.ExportAll;
            //spreadStreamExport.RunExport(@path, new SpreadStreamExportRenderer());
            
            Process p = new Process();
            p.StartInfo.FileName = path;
            p.Start();
            radWaiting.Visible = false;
            radWaiting.StartWaiting();
        }

        private void radBTNleft_Click(object sender, EventArgs e)
        {
            List<ListViewDataItem> itemsToRemove = new List<ListViewDataItem>();

            foreach (ListViewDataItem checkedItem in radListViewRigth.CheckedItems)
            {
                string columnName = checkedItem.Text;
                radListViewLeft.Items.Add(columnName);
                itemsToRemove.Add(checkedItem);
            }
            foreach (ListViewDataItem itemToRemove in itemsToRemove)
            {
                radListViewRigth.Items.Remove(itemToRemove);
            }
            GridView1.Columns.Clear();
            GridView1.Rows.Clear();
            
            FillAndupdateDataGridView();
        }    

       

       
    }
}
