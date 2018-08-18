using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.Remoting.Messaging;

namespace Name_Change_Data_Tool
{
    public class CSVItem
    {
        private string outDir;
        private string fileName;

        public CSVItem(string fName, string OutDir)
        {
            outDir = OutDir;
            fileName = outDir + "\\" + fName + DateTime.Now.ToString(" yyyy-MM-dd") + ".csv";
        }

        public void writeColumnHeaders()
        {
            string columns = "Alien Number,Certificate No,Applicant Name,Name Change,Date of Birth,Gender,Height,Marital,Country of Former Nationality,Residing At,Filing Location,Conducted By,Oath Ceremony Location,Date/Time of Oath Ceremony" + Environment.NewLine;
            File.WriteAllText(fileName, columns);
        }

        public void writeData(string data)
        {

            File.AppendAllText(fileName, data + Environment.NewLine);
        }
    }
}
