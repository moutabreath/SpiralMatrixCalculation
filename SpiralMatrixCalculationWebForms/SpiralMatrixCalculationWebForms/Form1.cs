using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Excel;

namespace SpiralMatrixCalculationWebForms
{

    public partial class Form1 : Form
    {

        private DataSet userMatrix;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
            //Workbook workBook = _excelApp.Workbooks.Open("myMatrix.xlsx");
            //Worksheet sheet = (Worksheet)workBook.Sheets[1];

            ////
            //// Take the used range of the sheet. Finally, get an object array of all
            //// of the cells in the sheet (their values). You can do things with those
            //// values. See notes about compatibility.
            ////
            //Range excelRange = sheet.UsedRange;
            //var numberOfColumns = excelRange.Columns.Count;
            //var numberOfRows = excelRange.Rows.Count;
            //object[,] valueArray = excelRange.get_Value(
            //    XlRangeValueDataType.xlRangeValueDefault);
            const int numberOfColumns = 4;
            const int numberOfRows = 4;
            int[,] valueArray = new int[numberOfRows, numberOfColumns] { { 1, 2, 3, 4 }, { 12, 13, 14, 5 }, { 11, 16, 15, 6 }, { 10, 9, 8, 7 } };
            

            CalculateMatrixSum(valueArray, numberOfColumns, numberOfRows);

        }

        private void CalculateMatrixSum(int[,] valueArray, int numberOfColumns, int numberOfRows)
        {
            int generalSum = 0;
            for (int i = 0; i < numberOfRows; i++)
            {
                generalSum += CalculateSingleRange(i, i, 1, numberOfColumns, valueArray, 1);
                generalSum += CalculateSingleRange(i - 1, 1, numberOfRows - i, i - 1, valueArray, 2);
                // i - 2 because we start from +2 index
                int column = numberOfColumns - (i - 2);
                generalSum += CalculateSingleRange(numberOfColumns - (i - 2), column, -1, i - 2, valueArray, 3);
                generalSum += CalculateSingleRange(numberOfRows - (i - 3), i - 3, -1, i - 3, valueArray, 4);

            }
            label1.Text = Convert.ToString(generalSum);
        }

        private int CalculateSingleRange(int rowNumber, int columnNumber, int step, int end, int[,] valueArray, int direction)
        {
            int miniSum = 0;
            if (direction == 1 || direction == 3)
            {
                while (rowNumber != end-1)
                {
                    miniSum += Convert.ToInt32(valueArray[rowNumber, columnNumber]);
                    rowNumber += step;
                }
            }
            else
            {
                while (columnNumber != end)
                {
                    try
                    {
                        miniSum += Convert.ToInt32(valueArray[rowNumber, columnNumber]);
                    }
                    catch (Exception e)
                    {
                        var msg = e.Message;
                    }
                    columnNumber += step;
                }
            }
            return miniSum;
        }

    }
}
