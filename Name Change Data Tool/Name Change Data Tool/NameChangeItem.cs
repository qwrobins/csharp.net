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
    public class NameChangeItem
    {
        private string File { get; set; }
        private int recCount = 0;
        private int maxVal = 0;
        private string status = "";
        private string parseResult = "";
        ArrayList dataList = new ArrayList();
        

        public NameChangeItem(string file)
        {
            File = file;
        }

        public string[] getNameChangeData()
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

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
            xlRange = xlWorksheet.UsedRange;

            int rCnt = xlRange.Rows.Count;
            int cCnt = xlRange.Columns.Count;
            int startRow = 0;

            status = "Gathering and analyzing name change data...";
            setStatusLabelInAnotherThread(status);

            for (int i = 1; i <= xlRange.Rows.Count; i++)
            {
                for (int j = 1; j <= xlRange.Columns.Count; j++)
                {
                    if ((((Excel.Range)xlWorksheet.Cells[i, j]).Value2 != null) && (((Excel.Range)xlWorksheet.Cells[i, j]).Value2.ToString() == "Alien Registration No."))
                    {
                        if (startRow == 0)
                        {
                            startRow = i + 1;
                            break;
                        }
                    }
                }

            }

            for (int i = startRow; i <= xlRange.Rows.Count; i++)
            {
                parseResult = string.Empty;
                for (int j = 2; j <= xlRange.Columns.Count; j++)
                {
                    if ((((Excel.Range)xlWorksheet.Cells[i, j]).Value2 != null) && (j <= 10))
                    {
                        parseResult += (string)((Excel.Range)xlWorksheet.Cells[i, j]).Value2.ToString() + "%";
                    }

                    if (j == 10)
                    {
                        if (((Excel.Range)xlWorksheet.Cells[(i + 1), 2]).Value2 == null)
                        {
                            if ((((Excel.Range)xlWorksheet.Cells[(i + 1), 4]).Value2 != null) && (((Excel.Range)xlWorksheet.Cells[(i + 1), 4]).Value2.ToString().Contains("*")))
                            {
                                parseResult += (string)((Excel.Range)xlWorksheet.Cells[(i + 1), 4]).Value2.ToString();
                            }
                        }
                    }
                }

                sortParseData(parseResult);
            }

            xlWorkbook.Close(false, Type.Missing, Type.Missing);
            xlApplication.Quit();
            GC.Collect();
            string[] nameChangeArray = (string[])dataList.ToArray(typeof(string));
            return nameChangeArray;
        }

        private void sortParseData(string data)
        {
            string[] dataArray = data.Split('%');
            int dc = dataArray.Count();
            string dataTextElement = "";
            if (dataArray.Count() > 2)
            {
                dataTextElement += dataArray[0] + ",";
                dataTextElement += dataArray[2] + ",";
                dataTextElement += "\"" + dataArray[1] + "\"" + ",";
                if (dataArray[3].ToString() != string.Empty)
                {
                    dataTextElement += "\"" + dataArray[3] + "\"" + ",";
                }
                else
                {
                    dataTextElement += ",";
                }

                dataList.Add(dataTextElement);
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
