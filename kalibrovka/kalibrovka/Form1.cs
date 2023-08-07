using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace kalibrovka
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            ShowIcon = false;

            InitializeComponent();

            ToolStripMenuItem fileItem = new ToolStripMenuItem("import Spectrum Files");

            ToolStripMenuItem fon = new ToolStripMenuItem("fon") { Checked = false, CheckOnClick = true };
            fon.Click += fon_Click;
            fileItem.DropDownItems.Add(fon);

            ToolStripMenuItem americium = new ToolStripMenuItem("americium") { Checked = false, CheckOnClick = true };
            americium.Click += americium_Click;
            fileItem.DropDownItems.Add(americium);

            ToolStripMenuItem radon = new ToolStripMenuItem("radon") { Checked = false, CheckOnClick = true };
            radon.Click += radon_Click;
            fileItem.DropDownItems.Add(radon);

            ToolStripMenuItem strontium = new ToolStripMenuItem("strontium") { Checked = false, CheckOnClick = true };
            strontium.Click += strontium_Click;
            fileItem.DropDownItems.Add(strontium);

            ToolStripMenuItem cesium = new ToolStripMenuItem("cesium") { Checked = false, CheckOnClick = true };
            cesium.Click += cesium_Click;
            fileItem.DropDownItems.Add(cesium);

            ToolStripMenuItem carbon = new ToolStripMenuItem("carbon") { Checked = false, CheckOnClick = true };
            carbon.Click += carbon_Click;
            fileItem.DropDownItems.Add(carbon); 

            menuStrip1.Items.Add(fileItem);
        }

        void fon_Click(object sender, EventArgs e)
        {
            chart();
        }

        void americium_Click(object sender, EventArgs e)
        {
            chart();
        }

        void radon_Click(object sender, EventArgs e)
        {
            chart();
        }

        void strontium_Click(object sender, EventArgs e)
        {
            chart();
        }

        void cesium_Click(object sender, EventArgs e)
        {
            chart();
        }

        void carbon_Click(object sender, EventArgs e)
        {
            chart();
        }

        

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //private void button1_Click(object sender, EventArgs e)
        //{

        //    string fname = "";
        //    OpenFileDialog fdlg = new OpenFileDialog();
        //    fdlg.Title = "Excel File Dialog";
        //    fdlg.InitialDirectory = @"c:\";
        //    fdlg.Filter = "Excel Files (*.xlsx; *.xls)|*.xlsx;*.xls";
        //    fdlg.RestoreDirectory = true;
        //    if (fdlg.ShowDialog() == DialogResult.OK)
        //    {
        //        fname = fdlg.FileName;

        //    }


        //    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        //    Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
        //    Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
        //    Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

        //    int rowCount = xlRange.Rows.Count;
        //    int colCount = xlRange.Columns.Count;

        //    // dt.Column = colCount;  
        //    dataGridView1.ColumnCount = colCount;
        //    dataGridView1.RowCount = rowCount;

        //    for (int i = 1; i <= rowCount; i++)
        //    {
        //        for (int j = 1; j <= colCount; j++)
        //        {


        //            //write the value to the Grid  


        //            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
        //            {
        //                dataGridView1.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();
        //            }
        //            // Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");  

        //            //add useful things here!     
        //        }
        //    }

        //    //cleanup  
        //    GC.Collect();
        //    GC.WaitForPendingFinalizers();

        //    //rule of thumb for releasing com objects:  
        //    //  never use two dots, all COM objects must be referenced and released individually  
        //    //  ex: [somthing].[something].[something] is bad  

        //    //release com objects to fully kill excel process from running in the background  
        //    Marshal.ReleaseComObject(xlRange);
        //    Marshal.ReleaseComObject(xlWorksheet);

        //    //close and release  
        //    xlWorkbook.Close();
        //    Marshal.ReleaseComObject(xlWorkbook);

        //    //quit and release  
        //    xlApp.Quit();
        //    Marshal.ReleaseComObject(xlApp);



        //}

        void chart()
        {
            // Создаем объект приложения Excel
            Application excelApp = new Application();

            // Открываем выбранный файл Excel
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xlsx; *.xls)|*.xlsx;*.xls";
            openFileDialog.InitialDirectory = @"C:\Users\misha\Desktop\Работа\064";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                Workbook workbook = excelApp.Workbooks.Open(openFileDialog.FileName);

                // Получаем первый лист
                Worksheet worksheet = (Worksheet)workbook.Sheets[1];

                // Получаем используемый диапазон ячеек на листе
                Range range = worksheet.UsedRange;

                // Получаем количество строк и столбцов
                int rowCount = range.Rows.Count;
                int columnCount = range.Columns.Count;

                // Создаем новую серию для данных из текущей таблицы
                var newSeries = new System.Windows.Forms.DataVisualization.Charting.Series();
                newSeries.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;

                // Добавляем точки из текущей таблицы на новую серию
                for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++)
                {
                    double xValue = rowIndex - 1;
                    double yValue = Convert.ToDouble(((Range)range.Cells[rowIndex, 2]).Text);
                    newSeries.Points.AddXY(xValue, yValue);
                }

                // Добавляем новую серию на график
                chart1.Series.Add(newSeries);

                // Закрываем приложение Excel
                workbook.Close();
                excelApp.Quit();

                // Устанавливаем заголовок графика
                chart1.Titles.Clear();
                chart1.Titles.Add("График данных из Excel файла");
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
