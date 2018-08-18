using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Name_Change_Data_Tool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private string fName;
        private string[] valData;
        private string[] nciData;
        private ArrayList finalList = new ArrayList();

        private void fOpenButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.AutoUpgradeEnabled = true;
            ofd.Filter = "Excel 97-2003 Files (*.xls)|*.xls|Excel 2007-2010 Files (*.xlsx)|*.xlsx";
            ofd.InitialDirectory = ".";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                fTextBox.Text = ofd.FileName;
                fName = ofd.SafeFileName;
            }
        }

        private void dirOpenButton_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Choose a directory to save the CSV file in";
                folderDialog.ShowNewFolderButton = true;
                folderDialog.RootFolder = Environment.SpecialFolder.MyComputer;
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string outFolder = folderDialog.SelectedPath;
                    dirTextBox.Text = outFolder;
                }
            }
        }

        private void convertButton_Click(object sender, EventArgs e)
        {
            if ((fTextBox.Text != "") && (dirTextBox.Text != ""))
            {
                bgw1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Please ensure that you have selected a file to open and a directory to write to before trying to convert");
            }
        }

        private void bgw1_DoWork(object sender, DoWorkEventArgs e)
        {
            NameChangeItem nci = new NameChangeItem(fTextBox.Text);
            nciData = nci.getNameChangeData();

            for (int i = 0; i < nciData.Length; i++)
            {
                if (nciData[i].ToString().Contains("*"))
                {
                    string pHolder = "";
                    string noStar = "";
                    pHolder = nciData[i].ToString();
                    noStar = pHolder.Replace("*", "");
                    nciData[i] = noStar;
                }
            }

            ValidationItem vi = new ValidationItem(fTextBox.Text);
            valData = vi.getValidationData();
        }

        private void bgw1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            statusLabel.Text = "Processing completed. Processing final record data and preparing to write to CSV file...";
            for (int i = 0; i < nciData.Length; i++)
            {
                finalList.Add(nciData[i] + valData[i]);
            }
            statusLabel.Text = "Writing CSV file to " + dirTextBox.Text;

            string fNameWithExt = fName;
            if (fNameWithExt.Contains("xlsx"))
            {
                fName = fNameWithExt.Replace(".xlsx", "");
            }
            else if (fNameWithExt.Contains(".xls"))
            {
                fName = fNameWithExt.Replace(".xls", "");
            }

            CSVItem csvi = new CSVItem(fName, dirTextBox.Text);
            csvi.writeColumnHeaders();
            string[] finalData = (string[])finalList.ToArray(typeof(string));
            for (int i = 0; i < finalData.Length; i++)
            {
                csvi.writeData(finalData[i]);
            }

            //parseTextBox.Text += "There are " + nciData.Length + " name change items and " + valData.Length + " validation items" + Environment.NewLine;
            //for (int i = 0; i < nciData.Length; i++)
            //{
            //    parseTextBox.Text += nciData[i].ToString() + Environment.NewLine;
            //}
            //for (int i = 0; i < valData.Length; i++)
            //{
            //    parseTextBox.Text += valData[i].ToString() + Environment.NewLine;
            //}
            statusLabel.Text = "";
            MessageBox.Show("Conversion completed successfully!");
        }

        internal void SetParseTextBox(string text)
        {
            parseTextBox.Text = text;
        }

        internal void setPbValue(Int32 pbValue)
        {
            if (pg1.InvokeRequired)
            {
                pg1.Invoke(new MethodInvoker(delegate()
                {
                    setPbValue(pbValue);
                }));
            }
            else
            {
                if (pbValue != pg1.Maximum)
                {
                    pg1.Value = pbValue;
                }
                else
                {
                    pg1.Value = pbValue;
                    pg1.Value = pg1.Minimum;
                    statusLabel.Text = "";
                }
            }
        }

        internal void setPbMaximum(Int32 pbMaxValue)
        {
            if (pg1.InvokeRequired)
            {
                pg1.Invoke(new MethodInvoker(delegate()
                    {
                        setPbMaximum(pbMaxValue);
                    }));
            }
            else
            {
                pg1.Maximum = pbMaxValue;
            }
        }

        internal void updateStatusLabel(String status)
        {
            if (statusLabel.InvokeRequired)
            {
                statusLabel.Invoke(new MethodInvoker(delegate()
                    {
                        updateStatusLabel(status);
                    }));
            }
            else
            {
                statusLabel.Text = status;
            }
        }

        internal void updateParseText(String parseText)
        {
            if (parseTextBox.InvokeRequired)
            {
                parseTextBox.Invoke(new MethodInvoker(delegate()
                    {
                        updateParseText(parseText);
                    }));
            }
            else
            {
                parseTextBox.Text = parseText;
            }
        }
    }
}
