using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.DataAccess.Client;
using System.Data;
using System.Configuration;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;


namespace TepProcess
{
    class TepProcess
    {
        static void Main(string[] args)
        {
            try
            {

            }
            catch (Exception ex)
            {
                string errMessage = Environment.NewLine + " " + DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt") + " " + ex.Message + " :: INNER - " + (ex.InnerException != null ? ex.InnerException.Message : string.Empty);
                File.AppendAllText(ConfigurationManager.AppSettings["ErrorPath"], errMessage);
                Console.WriteLine("Error Last : " + errMessage);
            }
        }

        // Reading Excel and Converting it to DataTable
        public class TepReadExcel
        {
            public string ReadTEPProgramExcel(string QueryName)
            {
                string excelPath = ConfigurationManager.AppSettings["QueryExcelPath"];
                DataTable dtResult = null;
                DataSet dsFull = new DataSet();
                int totalSheet = 0; //No of sheets on excel file  
                using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + excelPath + "';Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))
                {
                    objConn.Open();
                    OleDbCommand cmd = new OleDbCommand();
                    OleDbDataAdapter oleda = new OleDbDataAdapter();
                    DataSet ds = new DataSet();
                    DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string sheetName = string.Empty;
                    if (dt != null)
                    {
                        var tempDataTable = (from dataRow in dt.AsEnumerable()
                                             where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                             select dataRow).CopyToDataTable();
                        dt = tempDataTable;
                        totalSheet = dt.Rows.Count;
                        for (int i = 0; i < totalSheet; i++)
                        {
                            sheetName = dt.Rows[i]["TABLE_NAME"].ToString();

                            cmd.Connection = objConn;
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                            oleda = new OleDbDataAdapter(cmd);
                            oleda.Fill(ds, sheetName);
                            dtResult = ds.Tables[sheetName].Copy();
                        }
                    }
                    objConn.Close();




                    return dtResult; //Returning Dattable  
                }
            }
        }

        //Query Excecution with Query Name
        public DataTable TEPQueryExecution_Loop(string QueryName)
        {
            OracleConnection conn = null;
            OracleCommand cmd = null;
            try
            {
                Console.WriteLine(DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss") + " TEP query execution started");

                DataTable dtOutPut = new DataTable();
                dtOutPut.Columns.Add("USER_CONCURRENT_PROGRAM_NAME");
                dtOutPut.Columns.Add("PHASE_CODE");
                dtOutPut.Columns.Add("STATUS_CODE");
                dtOutPut.Columns.Add("PRIORITY");
                dtOutPut.Columns.Add("COMPLETION_TEXT");
                dtOutPut.Columns.Add("USER_NAME");
                dtOutPut.Columns.Add("HIGH_CRITICALITY");

                string connString = ConfigurationManager.AppSettings["ConnectionString"];

                #region Alert Query
                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        conn = new OracleConnection(connString);
                        conn.Open();
                        break;
                    }
                    catch (Exception ex)
                    {
                        string errMessage = Environment.NewLine + " " + DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt") + " " + ex.Message + " :: INNER - " + (ex.InnerException != null ? ex.InnerException.Message : string.Empty);
                        File.AppendAllText(ConfigurationManager.AppSettings["ErrorPath"], errMessage);
                        Console.WriteLine("Error DB Connection : " + errMessage);
                        Thread.Sleep(Convert.ToInt32(ConfigurationManager.AppSettings["DBConnectionRetryTime"]));
                    }
                }

                cmd = conn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = ConfigurationManager.AppSettings["TEPAlertQuery"];
                //cmd.Parameters.AddWithValue("param1", 1);
                cmd.ExecuteNonQuery();
                conn.Dispose();
                conn.Close();
                conn = null;
                cmd = null;
                #endregion

                conn = new OracleConnection(connString);
                conn.Open();
                string selectQuery = ConfigurationManager.AppSettings["TEPSelectQuery"];
                selectQuery = selectQuery.Replace("#START_DATE#", startDate.ToString("dd-MMM-yyyy HH:mm:ss"));
                selectQuery = selectQuery.Replace("#END_DATE#", endDate.ToString("dd-MMM-yyyy HH:mm:ss"));
                cmd = conn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = selectQuery;

                OracleDataReader dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        DataRow drResult = dtOutPut.NewRow();
                        drResult["USER_CONCURRENT_PROGRAM_NAME"] = dr["USER_CONCURRENT_PROGRAM_NAME"].ToString();
                        drResult["PHASE_CODE"] = dr["PHASE_CODE"].ToString();
                        drResult["STATUS_CODE"] = dr["STATUS_CODE"].ToString();
                        drResult["PRIORITY"] = dr["PRIORITY"].ToString();
                        drResult["COMPLETION_TEXT"] = dr["COMPLETION_TEXT"].ToString();
                        drResult["USER_NAME"] = dr["USER_NAME"].ToString();
                        dtOutPut.Rows.Add(drResult);
                    }
                    Console.WriteLine(DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss") + " Total Record - " + dtOutPut.Rows.Count);
                }
                else
                {
                    Console.WriteLine(DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss") + " No Rows Found");
                }

                return dtOutPut;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Dispose();
                conn.Close();
                cmd = null;
            }
        }

        //Write Query Output in a Excel
        public static void Excel_Create(string fileName, DataTable dt)
        {
            //Create excel app object
            Excel.Application xlSamp = new Microsoft.Office.Interop.Excel.Application();

            //Create a new excel book and sheet
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            //Then add a sample text into first cell
            xlWorkBook = xlSamp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            int k = 0;
            for (int i = 1; i <= dt.Columns.Count; i++)
            {
                xlWorkSheet.Cells[1, i] = dt.Columns[k].ToString();
                k++;
            }


            for (int i = 2; i <= dt.Rows.Count; i++)
            {
                for (int j = 1; j <= dt.Columns.Count; j++)
                {
                    xlWorkSheet.Cells[i, j] = dt.Rows[i - 1][j - 1].ToString();
                }
            }

            //Save the opened excel book to custom location
            string location = @"D:\" + fileName + ".xls";//Dont forget, you have to add to exist location
            xlWorkBook.SaveAs(location, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlSamp.Quit();

            //release Excel Object 
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSamp);
                xlSamp = null;
            }
            catch (Exception ex)
            {
                xlSamp = null;
                Console.Write("Error " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        //Sent Email using outlook
        public bool SendOutlookMail(DataTable dtResult, string strTime)
        {
            Outlook.Application app = new Outlook.Application();
            Outlook.MailItem mailItem = app.CreateItem(Outlook.OlItemType.olMailItem);
            //Outlook.Application app = new Outlook.Application();
            //Outlook.MailItem mailItem = app.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = "TEP Batch Monitoring Notification - " + strTime;
            mailItem.To = ConfigurationManager.AppSettings["OutlookToMails"];
            string body = string.Empty;
            if (dtResult != null && dtResult.Rows != null && dtResult.Rows.Count > 0)
            {
                StringBuilder strBody = new StringBuilder();

                strBody.Append(@"<!DOCTYPE html>
<html>
<head>
<style>
table {
    font-family: arial, sans-serif;
    border-collapse: collapse;
    font-size:12px;
}

td, th {
    border: 1px solid #dddddd;
    text-align: left;
    padding: 8px;
}

tr:nth-child(even) {
    background-color: #dddddd;
}
</style>
</head>
<body>");
                strBody.Append("Find the Tep Batch Monitor Query Result <br/><br/>");

                strBody.Append("<table id='example' class='display'><thead><tr><th>USER_CONCURRENT_PROGRAM_NAME</th><th>STATUS_CODE</th>");
                strBody.Append("<th>HIGH_CRITICALITY</th><th>PRIORITY</th><th>COMPLETION_TEXT</th><th>USER_NAME</th></tr></thead><tbody>");

                foreach (DataRow dtRow in dtResult.Rows)
                {
                    strBody.Append("<tr>");
                    strBody.Append("<td>" + dtRow.Field<string>("USER_CONCURRENT_PROGRAM_NAME") + "</td>");
                    strBody.Append("<td>" + dtRow.Field<string>("STATUS_CODE") + "</td>");
                    strBody.Append("<td>" + dtRow.Field<string>("HIGH_CRITICALITY") + "</td>");
                    strBody.Append("<td>" + dtRow.Field<string>("PRIORITY") + "</td>");
                    strBody.Append("<td>" + dtRow.Field<string>("COMPLETION_TEXT") + "</td>");
                    strBody.Append("<td>" + dtRow.Field<string>("USER_NAME") + "</td>");
                    strBody.Append("</tr>");
                }
                strBody.Append("</tbody></table></body></html>");

                body = strBody.ToString();
            }
            else
            {
                body = "Find the Tep Batch Monitor Query Result <br/><br/> No Rows found";
            }
            //mailItem.Body = "This is the message.";
            mailItem.HTMLBody = body;
            ((Outlook._MailItem)mailItem).Send();
            return true;
        }
    }
}