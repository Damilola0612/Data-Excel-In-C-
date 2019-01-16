using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
//using Microsoft.Office.Interop.Excel;
using System.Data;
using System.ComponentModel;

namespace BasicCodeSnippet
{
  public static void GetExcel<T>(List<T> data)
        {
            #region using a list of objects
            XLWorkbook workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Excel data");
            var headerCount = 0;
            var rowValue = 1;
            var colValue = 0;
            foreach (var obj in data)
            {//to create the header
                foreach (var property in obj.GetType().GetProperties())
                {
                    worksheet.Cell(1, ++headerCount).SetValue(property.Name);
                };
                //styling the header
                worksheet.Range(1, 1, 1, headerCount).Style.Fill.SetBackgroundColor(XLColor.Black).Font.SetFontColor(XLColor.White).Font.SetBold().Font.SetItalic().Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                break;
            };
            foreach (var obj in data)
            {//to create the data 
                rowValue++;
                colValue = 0;
                foreach (var property in obj.GetType().GetProperties())
                {
                    worksheet.Cell(rowValue, ++colValue).SetValue(property.GetValue(obj, null));
                }
            }
            //at this point, the rowValue would hold the last row number, colValue, the last column number and headCount the number of rows
            worksheet.Columns().AdjustToContents();
            workbook.SaveAs("Excel-data.xlsx");
            #endregion
        }
}
