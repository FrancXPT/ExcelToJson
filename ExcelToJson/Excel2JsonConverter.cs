using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Text.RegularExpressions;
using ExcelDataReader;
using System.IO;
using System.Data;

namespace Excel2Json
{
    public class Excel2JsonConverter
    {
        public string JsonPath = "";
        public string JsonFilename = "";
        public void ExcelFileToJson(string excelpath)
        {
            DataSet excelData = ExcelToDataSet(excelpath);
            string spreadSheetJson = "";

            for (int i = 0; i < excelData.Tables.Count; i++)
            {
                spreadSheetJson = SheetToJson(excelData, excelData.Tables[i].TableName);
                string fileName = excelData.Tables[i].TableName.Replace(" ", string.Empty);
                if(JsonPath == "")
                {
                    if(JsonFilename!= "")
                    System.IO.File.WriteAllText(JsonFilename + ".json", spreadSheetJson);

                }
                else
                    System.IO.File.WriteAllText(JsonPath, spreadSheetJson);


            }
        }

        public string SheetToJson(DataSet excelDataSet, string sheetName)
        {
            DataTable dataTable = excelDataSet.Tables[sheetName];
            return Newtonsoft.Json.JsonConvert.SerializeObject(dataTable);
        }

        public DataSet ExcelToDataSet(string filePath)
        {
            // Get the excel data reader with the excel data
            IExcelDataReader excelReader = GeteReader(filePath);

            if (excelReader == null)
            {
                return null;
            }

            DataSet data = new DataSet();
            do
            {
                // Get the DataTable from the current spreadsheet
                DataTable table = GetSheetData(excelReader);

                if (table != null)
                {
                    // Add the table to the data set
                    data.Tables.Add(table);
                }
            }
            while (excelReader.NextResult()); // Read the next sheet

            return data;
        }


        public DataTable GetSheetData(IExcelDataReader excelReader)
        {
            if (excelReader == null)
            {
                return null;
            }

            // Create the table with the spreadsheet name
            DataTable table = new DataTable(excelReader.Name);
            table.Clear();

            string value = null;
            bool rowIsEmpty;

            while (excelReader.Read())
            {
                DataRow row = table.NewRow();
                rowIsEmpty = true;
                for (int i = 0; i < excelReader.FieldCount; i++)
                {
                    // If the column is null and this is the first row, skip
                    // to next iteration (do not want to include empty columns)
                    if (excelReader.IsDBNull(i) &&
                        (excelReader.Depth == 1 || i > table.Columns.Count - 1))
                    {
                        continue;
                    }

                  
                     //value = excelReader.IsDBNull(i) ? "" : excelReader.GetString(i);


                    if (excelReader.GetFieldType(i).ToString() == "System.Double")
                    {
                        value = excelReader.IsDBNull(i) ? "" : excelReader.GetDouble(i).ToString();
                    }

                    if (excelReader.GetFieldType(i).ToString() == "System.Int")
                    {
                        value = excelReader.IsDBNull(i) ? "" : excelReader.GetInt32(i).ToString();
                    }

                    if (excelReader.GetFieldType(i).ToString() == "System.Bool")
                    {
                        value = excelReader.IsDBNull(i) ? "" : excelReader.GetBoolean(i).ToString();
                    }

                    if (excelReader.GetFieldType(i).ToString() == "System.DateTime")
                    {
                        value = excelReader.IsDBNull(i) ? "" : excelReader.GetDateTime(i).ToString();
                    }

                    if (excelReader.GetFieldType(i).ToString() == "System.TimeSpan")
                    {
                        value = excelReader.IsDBNull(i) ? "" : excelReader.GetDateTime(i).ToString();
                    }

                    if (excelReader.GetFieldType(i).ToString() == "System.String")
                    {
                        value = excelReader.IsDBNull(i) ? "" : excelReader.GetString(i).ToString();
                    }


                    if (excelReader.Depth == 0)
                    {
                        table.Columns.Add(value);
                    }
                    else 
                    {
                        row[table.Columns[i]] = value;
                    }

                    if (!string.IsNullOrEmpty(value))
                    {
                        rowIsEmpty = false;
                    }
                }

                if (excelReader.Depth != 1 && !rowIsEmpty)
                {
                    table.Rows.Add(row);
                }
            }

            return table;
        }

        public IExcelDataReader GeteReader(string filePath)
        {
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader eReader;

            Regex xlsRegex = new Regex(@"^(.*\.(xls$))");
            Regex xlsxRegex = new Regex(@"^(.*\.(xlsx$))");

            if (xlsRegex.IsMatch(filePath))
            {
                // Reading from *.xls)
                eReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else if (xlsxRegex.IsMatch(filePath))
            {
                // Reading from *.xlsx)
                eReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            else
            {
                stream.Close();
                return null;
            }

            return eReader;
        }


        public List<string> GetSheetNamesInFile(string filePath)
        {
            List<string> sheets = new List<string>();
            IExcelDataReader excelReader = GeteReader(filePath);

            if (excelReader == null)
            {
                return sheets;
            }

            do
            {
                sheets.Add(excelReader.Name);
            }
            while (excelReader.NextResult()); 

            return sheets;
        }
    }
}
