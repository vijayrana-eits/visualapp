using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelWithMultipleSheets.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
aaasasdasdasdda;
            return View();
        }

        public void ExportMaterialExcel()
        {
            DataSet ds = new DataSet();
            List<NewSampleClass> NewSampleClasslist = new List<NewSampleClass>();
            NewSampleClass newSampleClass = new NewSampleClass();
            newSampleClass.id = "1";
            newSampleClass.name = "Test";
            NewSampleClasslist.Add(newSampleClass);

            newSampleClass = new NewSampleClass();
            newSampleClass.id = "123456789123456789123";
           
            
            newSampleClass.name = "Text";
            NewSampleClasslist.Add(newSampleClass);
            newSampleClass = new NewSampleClass();

            newSampleClass.id = "2";
            newSampleClass.name = "Test2";
            NewSampleClasslist.Add(newSampleClass);
            System.Data.DataTable dt2 = new System.Data.DataTable("New1");
            dt2 = ListToDataTable(NewSampleClasslist);
            DataColumn Col = dt2.Columns.Add("M2 Field", System.Type.GetType("System.String"));
            Col.SetOrdinal(0);// to put the column in position 0;
            dt2.Columns["M2 Field"].DefaultValue = "";
            ds.Tables.Add(dt2);
            System.Data.DataTable dt1 = new System.Data.DataTable("New2");
            List<NewSampleClass> NewSampleClasslist1 = new List<NewSampleClass>();
            NewSampleClass newSampleClass1 = new NewSampleClass();
            newSampleClass1.name = "Test4";
            newSampleClass1.id = "4";
            NewSampleClasslist1.Add(newSampleClass1);

            newSampleClass1 = new NewSampleClass();
            newSampleClass1.name = "Test3";
            newSampleClass1.id = "3";
            NewSampleClasslist1.Add(newSampleClass1);

            newSampleClass1 = new NewSampleClass();
            newSampleClass1.name = "Test3";
            newSampleClass1.id = "5";
            NewSampleClasslist1.Add(newSampleClass1);
            dt1 = ListToDataTable(NewSampleClasslist1);
            ds.Tables.Add(dt1);
            DataColumn Col1 = dt1.Columns.Add("M2 Field", System.Type.GetType("System.String"));
            Col1.SetOrdinal(0);// to put the column in position 0;
            dt1.Columns["M2 Field"].DefaultValue = "";
            ExportExcel(ds);
            //return File(filecontent, ExcelContentType, "Material_Excel.xlsx");
        }
      
        #region Generate Excel

        public static System.Data.DataTable ListToDataTable<T>(List<T> data)
        {

            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            System.Data.DataTable dataTable = new System.Data.DataTable();
            try
            {
                for (int i = 0; i < properties.Count; i++)
                {
                    PropertyDescriptor property = properties[i];
                    if (i == 0)
                    {
                        dataTable.Columns.Add(property.Name, Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType);
                    }
                    else
                    {
                        dataTable.Columns.Add(property.Name, Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType);
                    }
                }

                object[] values = new object[properties.Count];
                foreach (T item in data)
                {
                    for (int i = 0; i < values.Length; i++)
                    {
                        values[i] = properties[i].GetValue(item);
                    }

                    dataTable.Rows.Add(values);
                }
            }
            catch (Exception ex)
            { }
            return dataTable;
        }

        public void ExportExcel(System.Data.DataSet ds)
        {
           

            using (ExcelPackage package = new ExcelPackage())
            {
                List<string> SheetNames = new List<string>();

                SheetNames.Add("Employee Details");

                SheetNames.Add("IndividualCustomer Details");

                SheetNames.Add("Contact Details");
                Application ExcelApp = new Application();
                Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook = null;
                
                Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet = null;
                ExcelWorkBook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Range cells = ExcelWorkBook.Worksheets[1].Cells;

                //cells.NumberFormat = "General";
                int[] colcount = new int[] { 21, 15, 6 };
                string siteid = "";
                try

                {

                    for (int i = 1; i < ds.Tables.Count; i++)

                     ExcelWorkBook.Worksheets.Add(); //Adding New sheet in Excel Workbook

                    for (int i = 0; i < ds.Tables.Count; i++)

                    {
                        siteid = "";
                        int r = 1; // Initialize Excel Row Start Position  = 1
                        ExcelWorkSheet = ExcelWorkBook.Worksheets[i + 1];
                        
                        //Writing Columns Name in Excel Sheet
                        int colval = colcount[i];

                        for (int col = 1; col <= ds.Tables[i].Columns.Count; col++)
                        {
                            var s = ds.Tables[i].Columns[col - 1].ColumnName;
                            
                            ExcelWorkSheet.Cells[r, col] = ds.Tables[i].Columns[col - 1].ColumnName;
                            // Column data type text
                            ((Range)ExcelWorkSheet.Cells[r, col]).EntireColumn.NumberFormat = "@";
                        }

                        r++;
                        for (int row = 0; row < ds.Tables[i].Rows.Count; row++) //r stands for ExcelRow and col for ExcelColumn
                        {
                                // Excel row and column start positions for writing Row=1 and Col=1
                                if (i == 0 && siteid != "" && siteid != ds.Tables[i].Rows[row][i+1].ToString())
                                {
                                    siteid = ds.Tables[i].Rows[row][i+1].ToString();
                                    Microsoft.Office.Interop.Excel.Range rng = (Microsoft.Office.Interop.Excel.Range)ExcelWorkSheet.Cells[r,row];
                                    Microsoft.Office.Interop.Excel.Range row1 = rng.EntireRow;
                                    row1.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, false);
                              
                                //    for (int col11 = 1; col11 <= ds.Tables[i].Columns.Count; col11++)
                                //    {
                                //        ExcelWorkSheet.Cells[r, col11] = "";

                                //}
                                row--;
                                    r++;
                                    continue;
                                }
                                else if (i == 1 && siteid != "" && siteid != ds.Tables[i].Rows[row][i+1].ToString())
                                {
                                    siteid = ds.Tables[i].Rows[row][i+1].ToString();
                                    Microsoft.Office.Interop.Excel.Range rng = (Microsoft.Office.Interop.Excel.Range)ExcelWorkSheet.Cells[r, row];
                                    Microsoft.Office.Interop.Excel.Range row1 = rng.EntireRow;
                                    row1.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, false);
                                //for (int col11 = 1; col11 <= ds.Tables[i].Columns.Count; col11++)
                                //{
                                //    ExcelWorkSheet.Cells[r, col11] = "";
                                //}
                                row--;
                                    r++;
                                    continue;
                                }
                                else if (i == 2 && siteid != "" && siteid != ds.Tables[i].Rows[row][i+1].ToString())
                                {
                                    siteid = ds.Tables[i].Rows[row][i+1].ToString();
                                    Microsoft.Office.Interop.Excel.Range rng = (Microsoft.Office.Interop.Excel.Range)ExcelWorkSheet.Cells[r, row];
                                    Microsoft.Office.Interop.Excel.Range row1 = rng.EntireRow;
                                    row1.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, false);
                                //for (int col11 = 1; col11 <= ds.Tables[i].Columns.Count; col11++)
                                //{
                                //    ExcelWorkSheet.Cells[r, col11] = "";
                                //}
                                row--;
                                    r++;
                                    continue;
                                }
                                for (int col = 1; col <= ds.Tables[i].Columns.Count; col++)
                                {
                                    if (siteid == "")
                                    {
                                        siteid = ds.Tables[i].Rows[row][i+1].ToString();
                                    }
                                    var ssss = ds.Tables[i].Rows[row][col - 1].ToString();
                                    ExcelWorkSheet.Cells[r, col] = ds.Tables[i].Rows[row][col - 1].ToString();                                    
                                }
                                r++;
                        }
                        ExcelWorkSheet.Name = SheetNames[i];//Renaming the ExcelSheets

                    }
                    
                    string FileName = "M2 Extract Design_" + DateTime.Now.Millisecond+DateTime.Now.Minute+DateTime.Now.Second+".xslx";
                    string FilePath = System.Web.HttpContext.Current.Server.MapPath("~/Temp/");
                    ExcelWorkBook.SaveAs(FilePath+FileName);                  
                    ExcelWorkBook.Close();

                    System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
                    response.ClearContent();
                    response.Clear();
                    response.ContentType = "Application/x-msexcel";
                    response.AppendHeader("Content-Disposition", "attachment; filename=" + FileName + ";");
                    response.TransmitFile(FilePath + FileName);
                    response.End();

                    // Deletes the file on server
                    if (System.IO.File.Exists(FilePath + FileName))
                    {
                        System.IO.File.Delete(FilePath + FileName);
                    }
                }
                catch (Exception exHandle)
                {
                    Console.WriteLine("Exception: " + exHandle.Message);
                    Console.ReadLine();
                }
                finally
                {
                    foreach (Process process in Process.GetProcessesByName("Excel"))
                        process.Kill();
                }
            }

            //return result;
        }
        

        #endregion
        public static string ExcelContentType
        {
            get
            { return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; }
        }
        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
    public class NewSampleClass
    {
        public string id { get; set; }
        public string name { get; set; }
    }

}
