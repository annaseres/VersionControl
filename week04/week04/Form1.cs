using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace week04
{
    public partial class Form1 : Form
    {

        Excel.Application xlApp;
        Excel.Workbook xlWB;
        Excel.Worksheet xlSheet;


        RealEstateEntities context = new RealEstateEntities();
        List<Flat> Flats;

        public Form1()
        {
            InitializeComponent();
            LoadData();

        }

        private void LoadData()
        {
            Flats = context.Flat.ToList();
            
        }

        public void CreateExcel()
        {
            try
            {

                xlApp = new Excel.Application();

                xlWB = xlApp.Workbooks.Add(Missing.Value);

                xlSheet = xlWB.ActiveSheet;

                
                //CreateTable(); 

                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex) 
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }

        public void CreateTable()
        {
            string[] headers = new string[] {
             "Kód",
             "Eladó",
             "Oldal",
             "Kerület",
             "Lift",
             "Szobák száma",
             "Alapterület (m2)",
             "Ár (mFt)",
             "Négyzetméter ár (Ft/m2)"};

            xlSheet.Cells[1, 1] = headers[0];

            object[,] values = new object[Flats.Count, headers.Length];

            string lift = "";

            xlSheet.get_Range(
                GetCell(2, 1),
                GetCell(1 + values.GetLength(0), values.GetLength(1))).Value2 = values;
            

            int counter = 0;
            foreach (var item in Flats)
            {
                


                if (item.Elevator==true)
                {
                    
                    lift = "Van";
                }
                else
                {
                    lift = "Nincs";
                }

                values[counter, 0] = item.Code;
                values[counter, 1] = item.Vendor;
                values[counter, 2] = item.Side;
                values[counter, 3] = item.District;
                values[counter, 4] = lift;
                values[counter, 5] = item.NumberOfRooms;
                values[counter, 6] = item.FloorArea;
                values[counter, 7] = item.Price;
                values[counter, 8] = item.Price / item.FloorArea;
                counter++;


                Excel.Range headerRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1, headers.Length));
                headerRange.Font.Bold = true;
                headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                headerRange.EntireColumn.AutoFit();
                headerRange.RowHeight = 40;
                headerRange.Interior.Color = Color.LightBlue;
                headerRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

                int lastRowID = xlSheet.UsedRange.Rows.Count;
                Excel.Range wholetable=xlSheet.get_Range(GetCell(2,1), GetCell(lastRowID, headers.Length));
                wholetable.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
                wholetable.Font.Bold = true;
                wholetable.Interior.Color = Color.LightYellow;

                Excel.Range utolso = xlSheet.get_Range(GetCell(2, headers.Length), GetCell(lastRowID, headers.Length));
                utolso.Interior.Color = Color.LightGreen;

            }
        }


        private string GetCell(int x, int y)
        {
            string ExcelCoordinate = "";
            int dividend = y;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                ExcelCoordinate = Convert.ToChar(65 + modulo).ToString() + ExcelCoordinate;
                dividend = (int)((dividend - modulo) / 26);
            }
            ExcelCoordinate += x.ToString();

            return ExcelCoordinate;

            
        }

        

        
    }
}
