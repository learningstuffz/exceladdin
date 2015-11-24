using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Xml.Linq;
using System.Data;

namespace ExcelAddIn
{
    public partial class RibbonFetch
    {
        Excel.Range newFirstRow;
        Excel.Range firstRow;
        Excel.Worksheet activeWorksheet;
        private void RibbonFetch_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void btnFetch_Click(object sender, RibbonControlEventArgs e)
        {
            activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            firstRow = activeWorksheet.get_Range("A1");
            //firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            newFirstRow = activeWorksheet.get_Range("A1");

            try
            {
                firstRow.Value2 = "Fetching Data Please Wait...";
                string url = "http://localhost:5000/sample/getdata";
                WebClient wc = new WebClient();
                wc.DownloadStringCompleted += HttpsCompleted;
                wc.DownloadStringAsync(new Uri(url));

                

            }
            catch(Exception ex)
            {
                firstRow.Value2 = "An Error occured while accessing the service";
            }
            
        }
        private void HttpsCompleted(object sender, DownloadStringCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                string response = e.Result;
                DataTable xmlData = Newtonsoft.Json.JsonConvert.DeserializeObject<DataTable>(response);
                //DataTable dtObj = new DataTable("MyExcel");
                foreach (DataColumn cols in xmlData.Columns)
                {
                    newFirstRow.Value2 = cols.ColumnName;
                    newFirstRow = newFirstRow.Next;
                }
                newFirstRow = activeWorksheet.Cells[firstRow.Row + 1, firstRow.Column];
                object[] data = new object[xmlData.Columns.Count];
                foreach (DataRow dr in xmlData.Rows)
                {
                    data = new object[xmlData.Columns.Count];
                    data = dr.ItemArray;

                    //= (Excel.Range)activeWorksheet.Cells[newFirstRow.Row, firstRow.Column]; ;
                    Excel.Range newRangeData = newFirstRow.get_Resize(1, xmlData.Columns.Count);

                    newRangeData.set_Value(Type.Missing, data);
                    newFirstRow = activeWorksheet.Cells[newFirstRow.Row + 1, firstRow.Column];
                }
                firstRow.get_Offset(1, 1).get_Resize(xmlData.Rows.Count, xmlData.Columns.Count).Style = "Currency";
            }
            else
            {
                firstRow.Value2 = "An Error occured while accessing the service";
            }
        }
    }
}
