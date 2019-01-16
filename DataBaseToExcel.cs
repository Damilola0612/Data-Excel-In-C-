using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.ComponentModel;

namespace Excel
{
  public static void GetExcel<T>(List<T> data)
  {//depends on ClosedXML and works fine✔
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
  public static void GetExceldt<T>(List<T> data)
  {
      //convert to dataTable
      DataTable table = ToDataTable(data);
      XLWorkbook workbook = new XLWorkbook();
      var worksheet = workbook.AddWorksheet("Excel dataTable");
      for (var i = 1; i <= table.Columns.Count; i++)
      {
          worksheet.Cell(1, i).SetValue(table.Columns[i - 1].ColumnName);
      }
      worksheet.Range(1, 1, 1, table.Columns.Count).Style.Fill.SetBackgroundColor(XLColor.Black).Font.SetFontColor(XLColor.White).Font.SetBold().Font.SetItalic().Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
      var rowValue = 1;
      foreach (DataRow datarow in table.Rows)
      {
          rowValue++;
          for (var i = 1; i <= table.Columns.Count; i++)
          {
              worksheet.Cell(rowValue, i).SetValue(datarow[i - 1].ToString());
          }
      }
      worksheet.Columns().AdjustToContents();
      workbook.SaveAs("Excel-adjusted-dataTable.xlsx");
    //This function depends on the private method "ToDataTable" given below😊😊
  }//in addition, I can add option for themes
  private static DataTable ToDataTable<t>(IList<t> data)
  {
      PropertyDescriptorCollection properties =
          TypeDescriptor.GetProperties(typeof(t));
      DataTable table = new DataTable();
      foreach (PropertyDescriptor prop in properties)
          table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
      foreach (t item in data)
      {
          DataRow row = table.NewRow();
          foreach (PropertyDescriptor prop in properties)
              row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
          table.Rows.Add(row);
      }
      return table;
  }
  ///////////////////////////////////////////////////////////////////////////////////////////////////////
  ///////////////////////////////////////////////////////////////////////////////////////////////////////
  ///////////////////////////////////////////////////////////////////////////////////////////////////////
  //the method given below depends on "Microsoft.Office.Interop.Excel.dll" and would NOT work if you don't have excel installed
  public void GetSampleExcel()
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);

                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = "StudentReportCard";

                worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[1, 8]].Merge();
                worKsheeT.Cells[1, 1] = "Student Report Card";
                worKsheeT.Cells.Font.Size = 15;

                int rowcount = 2;

                foreach (DataRow datarow in ExportToExcel().Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= ExportToExcel().Columns.Count; i++)
                    {

                        if (rowcount == 3)
                        {
                            worKsheeT.Cells[2, i] = ExportToExcel().Columns[i - 1].ColumnName;
                            worKsheeT.Cells.Font.Color = System.Drawing.Color.Black;

                        }

                        worKsheeT.Cells[rowcount, i] = datarow[i - 1].ToString();


                        if (rowcount > 3)
                        {
                            if (i == ExportToExcel().Columns.Count)
                            {
                                if (rowcount % 2 == 0)
                                {
                                    celLrangE = worKsheeT.Range[worKsheeT.Cells[rowcount, 1], worKsheeT.Cells[rowcount, ExportToExcel().Columns.Count]];
                                }

                            }
                        }

                    }

                }
                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[rowcount, ExportToExcel().Columns.Count]];
                celLrangE.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;

                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[2, ExportToExcel().Columns.Count]];

                worKbooK.SaveAs(@"c:\TestASPdotNET.xlsx"); ;
                worKbooK.Close();
                excel.Quit();

            }
            catch (Exception ex)
            {

            }
            finally
            {
                worKsheeT = null;
                celLrangE = null;
                worKbooK = null;
            }

        }
        private System.Data.DataTable ExportToExcel()
        {
            System.Data.DataTable table = new System.Data.DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Sex", typeof(string));
            table.Columns.Add("Subject1", typeof(int));
            table.Columns.Add("Subject2", typeof(int));
            table.Columns.Add("Subject3", typeof(int));
            table.Columns.Add("Subject4", typeof(int));
            table.Columns.Add("Subject5", typeof(int));
            table.Columns.Add("Subject6", typeof(int));
            table.Rows.Add(1, "Amar", "M", 78, 59, 72, 95, 83, 77);
            table.Rows.Add(2, "Mohit", "M", 76, 65, 85, 87, 72, 90);
            table.Rows.Add(3, "Garima", "F", 77, 73, 83, 64, 86, 63);
            table.Rows.Add(4, "jyoti", "F", 55, 77, 85, 69, 70, 86);
            table.Rows.Add(5, "Avinash", "M", 87, 73, 69, 75, 67, 81);
            table.Rows.Add(6, "Devesh", "M", 92, 87, 78, 73, 75, 72);
            return table;
        }
}
