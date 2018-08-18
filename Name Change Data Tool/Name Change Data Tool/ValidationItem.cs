using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Runtime.Remoting.Messaging;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Name_Change_Data_Tool
{
    public class ValidationItem
    {
        private string File { get; set; }
        private int recCount = 0;
        private int maxVal = 0;
        private string status = "";
        string parseResult;
        private ArrayList dataList = new ArrayList();

        public ValidationItem(string file)
        {
            File = file;
        }

        public string[] getValidationData()
        {
            Excel.Application xlApplication;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            Excel.Range xlRange;
            xlApplication = new Excel.Application();
            xlWorkbook = xlApplication.Workbooks.Open(File, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            try
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(2);
                
            }
            catch
            {
                MessageBox.Show("The current file is missing a required worksheet. Please check the file and ensure that both worksheets are present. The file cannot be processed until this issue has been resolved");

                xlWorkbook.Close(false, Type.Missing, Type.Missing);
                xlApplication.Quit();
                GC.Collect();

                foreach (Process proc in Process.GetProcessesByName("EXCEL"))
                {
                    proc.Kill();
                }

                Application.Exit();
            }

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(2);
            xlRange = xlWorksheet.UsedRange;

            List<string> recList = new List<string>{};
            int rCnt = xlRange.Rows.Count;
            int cCnt = xlRange.Columns.Count;
            int startRow = 0;

            //MessageBox.Show("Found worksheet " + k);
            status = "Gathering and analyzing validation data...";
            setStatusLabelInAnotherThread(status);

            for (int i = 1; i <= xlRange.Rows.Count; i++)
            {
                for (int j = 1; j <= xlRange.Columns.Count; j++)
                {
                    if ((((Excel.Range)xlWorksheet.Cells[i, j]).Value2 != null) && (((Excel.Range)xlWorksheet.Cells[i, j]).Value2.ToString() == "Date of Birth"))
                    {
                        if (startRow == 0)
                        {
                            startRow = i;
                        }

                        maxVal++;
                        setPbMaximumInAnotherThread(maxVal);
                    }
                }
            }

            status = maxVal.ToString() + " records found";
            setStatusLabelInAnotherThread(status);

            for (int i = startRow; i <= xlRange.Rows.Count; i++)
            {
                for (int j = 1; j <= xlRange.Columns.Count; j++)
                {
                    if ((((Excel.Range)xlWorksheet.Cells[i, j]).Value2 != null) && (((Excel.Range)xlWorksheet.Cells[i, j]).Value2.ToString() == "Date of Birth"))
                    {
                        if (recCount == 0)
                        {
                            status = "Processing " + maxVal.ToString() + " records";
                            setStatusLabelInAnotherThread(status);
                        }

                        recCount++;
                        setPg1InAnotherThread(recCount);
                    }

                    if ((((Excel.Range)xlWorksheet.Cells[i, j]).Value2 != null) && (((Excel.Range)xlWorksheet.Cells[i, j]).Value2.ToString() != ":"))
                    {
                        parseResult += (string)((Excel.Range)xlWorksheet.Cells[i, j]).Value2.ToString() + "%";
                        removeLabels(parseResult);
                        if (parseResult != "")
                        {
                            sortParseData(parseResult);
                        }
                    }
                }
            }

            xlWorkbook.Close(false, Type.Missing, Type.Missing);
            xlApplication.Quit();
            GC.Collect();

            foreach (Process proc in Process.GetProcessesByName("EXCEL"))
            {
                proc.Kill();
            }

            string[] validationArray = (string[])dataList.ToArray(typeof(string));
            return validationArray;
        }

        private void removeLabels(string strVal)
        {
            strVal = strVal.Replace("Date of Birth%", string.Empty).Replace("Alien Number%", string.Empty).Replace("Gender%", string.Empty).Replace("Filing Location%", string.Empty)
                .Replace("Height%", string.Empty).Replace("Applicant%", string.Empty).Replace("Marital%", string.Empty).Replace("Conducted By%", string.Empty).Replace("Country of Former Nationality%", string.Empty)
                .Replace("Oath Ceremony Location%", string.Empty).Replace("Residing At:%", string.Empty).Replace("Date/Time of Oath Ceremony%", string.Empty).Replace(",,", ",").Replace("\r\n,", "\r\n").Replace("Number of s Scheduled to Appear at Oath Ceremony:", string.Empty);
            parseResult = strVal;
        }

        private void sortParseData(string data)
        {
            string[] dataArray = data.Split('%');
            string dataTextElement = "";
            int dc = dataArray.Count();
            if (dataArray.Count() == 13)
            {
                dataTextElement += dataArray[0] + ",";
                dataTextElement += dataArray[2] + ",";
                dataTextElement += dataArray[4] + ",";
                dataTextElement += dataArray[6] + ",";
                dataTextElement += "\"" + dataArray[8] + "\"" + ",";
                dataTextElement += "\"" + dataArray[10] + "\"" + ",";
                dataTextElement += "\"" + dataArray[3] + "\"" + ",";
                dataTextElement += dataArray[7] + ",";
                dataTextElement += "\"" + dataArray[9] + "\"" + ",";

                DateTime dt = Convert.ToDateTime(dataArray[11]);
                string convertedDate = dt.ToShortDateString();
                string convertedTime = dt.ToShortTimeString();
                dataArray[11] = convertedDate + " " + convertedTime;
                dataTextElement += dataArray[11];
                dataList.Add(dataTextElement);
                parseResult = string.Empty;
            }
        }

        public void setPg1InAnotherThread(Int32 val)
        {
            new Func<Int32>(setPbValue).BeginInvoke(new AsyncCallback(setPbValueCallback), null);
        }

        private Int32 setPbValue()
        {
            Int32 result = recCount;
            return result;
        }

        private void setPbValueCallback(IAsyncResult ar)
        {
            AsyncResult result = (AsyncResult)ar;
            Func<Int32> del = (Func<Int32>)result.AsyncDelegate;
            try
            {
                Int32 pbValue = del.EndInvoke(ar);
                if (pbValue != 0)
                {
                    Form1 frm1 = (Form1)findOpenForm(typeof(Form1));
                    if (frm1 != null)
                    {
                        frm1.setPbValue(pbValue);
                    }
                }
            }
            catch { }
        }

        public void setPbMaximumInAnotherThread(Int32 val)
        {
            new Func<Int32>(setPbMaxValue).BeginInvoke(new AsyncCallback(setPbMaxValueCallback), null);
        }

        private Int32 setPbMaxValue()
        {
            Int32 result = maxVal;
            return result;
        }

        private void setPbMaxValueCallback(IAsyncResult ar)
        {
            AsyncResult result = (AsyncResult)ar;
            Func<Int32> del = (Func<Int32>)result.AsyncDelegate;
            try
            {
                Int32 pbMaxValue = del.EndInvoke(ar);
                if (pbMaxValue > 0)
                {
                    Form1 frm1 = (Form1)findOpenForm(typeof(Form1));
                    if (frm1 != null)
                    {
                        frm1.setPbMaximum(pbMaxValue);
                    }
                }
            }
            catch { }
        }

        public void setStatusLabelInAnotherThread(String prStatus)
        {
            new Func<String>(setStatus).BeginInvoke(new AsyncCallback(setStatusLabelCallback), null);
        }

        private String setStatus()
        {
            String result = status;
            return result;
        }

        private void setStatusLabelCallback(IAsyncResult ar)
        {
            AsyncResult result = (AsyncResult)ar;
            Func<String> del = (Func<String>)result.AsyncDelegate;
            try
            {
                String prStatus = del.EndInvoke(ar);
                if (prStatus != "")
                {
                    Form1 frm1 = (Form1)findOpenForm(typeof(Form1));
                    if (frm1 != null)
                    {
                        frm1.updateStatusLabel(prStatus);
                    }
                }
            }
            catch { }
        }

        public void setParseTextInAnotherThread(String parseResult)
        {
            new Func<String>(setParseResult).BeginInvoke(new AsyncCallback(setParseTextCallback), null);
        }

        private String setParseResult()
        {
            String result = parseResult;
            return result;
        }

        private void setParseTextCallback(IAsyncResult ar)
        {
            AsyncResult result = (AsyncResult)ar;
            Func<String> del = (Func<String>)result.AsyncDelegate;
            try
            {
                String parseText = del.EndInvoke(ar);
                if (parseText != "")
                {
                    Form1 frm1 = (Form1)findOpenForm(typeof(Form1));
                    if (frm1 != null)
                    {
                        frm1.updateParseText(parseText);
                    }
                }
            }
            catch { }
        }

        private static Form findOpenForm(Type typ)
        {
            for (int i = 0; i < Application.OpenForms.Count; i++)
            {
                if (!Application.OpenForms[i].IsDisposed && (Application.OpenForms[i].GetType() == typ))
                {
                    return Application.OpenForms[i];
                }
            }
            return null;
        }
    }
}
