//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Configuration;

namespace SFDCtoD365ExcelWrapper
{
    class Program
    {
       
        /// <summary>
        /// NOtes - 
        /// Set Initial row for reading
        /// Set Initial column for compare fields
        /// </summary>
        private static int rowCount = 0;
        private static int colCount = 0;

        static void Main(string[] args)
        {
            DataSet ds = TryDatasetforSpecificWorksheet();

            string salesforceExcelFilter = ConfigurationManager.AppSettings["SalesforceExcelFilter"];
            ds.Tables[0].DefaultView.RowFilter = salesforceExcelFilter;
            DataTable dt = (ds.Tables[0].DefaultView).ToTable();

            UpdateExcelForDynamics(dt);

        }

        private static void UpdateExcelForDynamics(DataTable dt)
        {
            //Instance reference for Excel Application
            Microsoft.Office.Interop.Excel.Application objXL = null;
            //Workbook refrence
            Microsoft.Office.Interop.Excel.Workbook objWB = null;

            DataSet ds = new DataSet();
            try
            {
                string dynamicsFilePath = @ConfigurationManager.AppSettings["DynamicsExcelLocation"];
            objXL = new Microsoft.Office.Interop.Excel.Application();
            objWB = objXL.Workbooks.Open(dynamicsFilePath);//Your path to excel file.

            // Only retrieve first worksheet with new fields
            Microsoft.Office.Interop.Excel.Worksheet NewFieldsWorksheet = objWB.Worksheets[1];

            // setting range for second row
            Excel.Range xlRange = NewFieldsWorksheet.UsedRange;
            int rowNumber = xlRange.Rows.Count + 1;

            string prefix = ConfigurationManager.AppSettings["Prefix"];
            string Entity = ConfigurationManager.AppSettings["EntityName"];

            foreach (DataRow dtRow in dt.Rows)
            {
                //Common fields
                NewFieldsWorksheet.Cells[rowNumber, 2] = dtRow["Label"].ToString(); // Label Name
                NewFieldsWorksheet.Cells[rowNumber, 3] = GetSchemaWithPrefix(dtRow["Label"].ToString());// schema
                NewFieldsWorksheet.Cells[rowNumber, 4] = GetDynamicsType(dtRow["Type"].ToString()); // field type
                NewFieldsWorksheet.Cells[rowNumber, 5] = Entity; // entity
                NewFieldsWorksheet.Cells[rowNumber, 7] = GetDynamicsRequired(dtRow["IsRequired"].ToString()); // field type

                // Special Set Up based on Type
                if (dtRow["Type"].ToString() == "STRING")
                {
                    NewFieldsWorksheet.Cells[rowNumber, 12] = dtRow["Length"].ToString(); // If text - Lenght
                    NewFieldsWorksheet.Cells[rowNumber, 13] = "Text";
                }

                rowNumber++;
            }

            //// Disable file override confirmaton message  
            objXL.DisplayAlerts = false;
            objWB.SaveAs(dynamicsFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
                    Excel.XlSaveConflictResolution.xlLocalSessionChanges, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value);
            objWB.Close();
            objXL.Quit();

            Marshal.ReleaseComObject(NewFieldsWorksheet);
                Marshal.ReleaseComObject(objWB);
                Marshal.ReleaseComObject(objXL);

            }

            catch (Exception ex)
            {
                objWB.Saved = true;
                //Closing work book
                objWB.Close();
                //Closing excel application
                objXL.Quit();

            }

        }

        private static string GetDynamicsRequired(string isRequired)
        {
            if (isRequired == "TRUE")
                return "System required";

            return "";
        }

        private static string GetSchemaWithPrefix(string name)
        {
            string prefix = "new_";
            return prefix + name.Replace(" ", "");
        }

        private static string GetDynamicsType(string Salestype)
        {
            if (Salestype == "STRING")
                return "Single line of text";

            return "";
        }

        public static DataSet TryDatasetforSpecificWorksheet()
        {
            //Instance reference for Excel Application
            Microsoft.Office.Interop.Excel.Application objXL = null;
            //Workbook refrence
            Microsoft.Office.Interop.Excel.Workbook objWB = null;
            DataSet ds = new DataSet();
            try
            {
                objXL = new Microsoft.Office.Interop.Excel.Application();

                string SalesforceFilePath = @ConfigurationManager.AppSettings["SalesforceExcelLocation"];
                objWB = objXL.Workbooks.Open(SalesforceFilePath);//Your path to excel file.

                // Only retrieve second worksheet with mapping information from field dump
                Microsoft.Office.Interop.Excel.Worksheet objSHT = objWB.Worksheets[2];

                //foreach (Microsoft.Office.Interop.Excel.Worksheet objSHT in objWB.Worksheets)
                //{
                int rows = objSHT.UsedRange.Rows.Count;
                int cols = objSHT.UsedRange.Columns.Count;

                DataTable dt = new DataTable();
                int noofrow = 3;
                //If 1st Row Contains unique Headers for datatable include this part else remove it
                //Start
                for (int c = 1; c <= cols; c++)
                {
                    int rowForHeaders = 3; // Marking columns headers
                    string colname = objSHT.Cells[rowForHeaders, c].Text;
                    string trimcolname = colname.Replace(" ", "");

                    dt.Columns.Add(trimcolname);
                    noofrow = rowForHeaders + 1;// Mark new row # in order to get information
                }
                //END
                for (int r = noofrow; r <= rows; r++)
                {
                    DataRow dr = dt.NewRow();
                    for (int c = 1; c <= cols; c++)
                    {
                        dr[c - 1] = objSHT.Cells[r, c].Text;
                    }
                    dt.Rows.Add(dr);
                }
                ds.Tables.Add(dt);
                //}
                //Closing workbook
                objWB.Close();
                //Closing excel application
                objXL.Quit();
                return ds;
            }

            catch (Exception ex)
            {
                objWB.Saved = true;
                //Closing work book
                objWB.Close();
                //Closing excel application
                objXL.Quit();
                //Response.Write("Illegal permission");
                return ds;
            }

        }

    }
}
