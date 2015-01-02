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
            
            Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workBook = _excelApp.Workbooks.Open(@"C:\Users\Tal\Documents\myMatrix.xlsx");
            Worksheet sheet = (Worksheet)workBook.Sheets[1];

            //
            // Take the used range of the sheet. Finally, get an object array of all
            // of the cells in the sheet (their values). You can do things with those
            // values. See notes about compatibility.
            //
            Range excelRange = sheet.UsedRange;
            var numberOfColumns = excelRange.Columns.Count;
            var numberOfRows = excelRange.Rows.Count;
            object[,] valueArray = excelRange.get_Value(
                XlRangeValueDataType.xlRangeValueDefault);
            CalculateMatrixSum(valueArray, numberOfColumns, numberOfRows);

        }

        private void CalculateMatrixSum(object[,] valueArray, int numberOfColumns, int numberOfRows)
        {
            int generalSum = 0;
            // Iterate through the columns. Since this is a matrix so length =  width, use Length as limit.
            for (int i = 1; i <= numberOfRows; i++)
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
        }

        private int CalculateSingleRange(int rowNumber, int columnNumber, int step, int end, object[,] valueArray, int direction)
        {
            int miniSum = 0;
            if (direction == 1 || direction == 3)
            {
                while (rowNumber != end)
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
