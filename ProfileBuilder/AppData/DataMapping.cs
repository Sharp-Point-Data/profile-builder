using System;
using System.Collections.Generic;
using System.Text;
using ExcelDataReader;
using System.Data;
using System.IO;
using System.Collections;

namespace ProfileBuilder.ExcelData
{
    class DataReader
    {
        public string ImageFileName { get; private set; }
        public string Name { get; private set; }
        public string JobTitle { get; private set; }
        public string ContactInfo { get; private set; }
        public string Bio { get; private set; }

        private DataSet dataSet;

        private int position = -1;
        private int end;



        public DataReader(string path)
        {
            FileInfo filePath = new FileInfo(path);

            FileStream stream = new FileStream(filePath.FullName, FileMode.Open, FileAccess.Read);
            
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(stream))
            {
                dataSet = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });

                position = -1;
                end = dataSet.Tables[0].Rows.Count - 1;
            }
        }

        public void readLine()
        {            
            DataRow row = dataSet.Tables[0].Rows[position];

            ImageFileName = row[0].ToString();
            Name = row[1].ToString();
            JobTitle = row[2].ToString();
            ContactInfo = row[3].ToString();
            Bio = row[4].ToString();            
        }

        public bool nextLine()
        {
            if (position < end)
            {
                ++position;
                return true;
            }
            else
                return false;
        }

        public override string ToString()
        {
            return "Name: " + Name + " | Job Title: " + JobTitle + " | Contact Info: " + ContactInfo + " | Bio : " + Bio;
        }
    }
}

          

