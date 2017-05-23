using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.IO;
using Excel;

using System.Windows.Forms;

namespace xls2gdb
{
    class ExcelReader
    {
        private string name = string.Empty;
        private bool flagHeaders = false;
        private FileStream stream;

        private List<string> worksheet = new List<string>();
        private List<string> fieldsheet = new List<string>();
        
        private DataSet dataFile = new DataSet();
        private IExcelDataReader excelReader = null;

        public void setSourceData(string fileName, bool flag)
        {
            name = fileName; flagHeaders = flag;
            if (excelReader != null)
                excelReader.IsFirstRowAsColumnNames = flagHeaders;
        }

        public bool openExcelReader()
        {
            try
            {
                if (name.Length != 0)
                {
                    stream = File.Open(name, FileMode.Open, FileAccess.Read);
                    excelReader = Path.GetExtension(name).ToLower() == ".xls" ?
                    ExcelReaderFactory.CreateBinaryReader(stream) :
                    ExcelReaderFactory.CreateOpenXmlReader(stream);
                    excelReader.IsFirstRowAsColumnNames = flagHeaders;
                }
                return !excelReader.IsClosed;
            }
            catch (Exception Except)
            {
                MessageBox.Show(Except.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            
        }
        public bool cloceExcelReader()
        {
            if (excelReader != null && !excelReader.IsClosed)
            {
                excelReader.Close();
                bool f = excelReader.IsClosed;
                excelReader.Dispose();
                return f;
            }
            return false;
        }
        public bool resetConnection()
        {
            cloceExcelReader();  
            return openExcelReader();
        }

        public int countSheets()
        {
            return excelReader.ResultsCount;
        }
        public void loadSheetsName()
        {
            worksheet.Clear();
            resetConnection();

            for (int i = 0; i < countSheets(); ++i)
            {
                selectSheet(i);
                worksheet.Add(excelReader.Name);
            }
            cloceExcelReader();
        }
        public List<string> worksheetNameList()
        {
            return worksheet;
        }
        private void selectSheet(int index)
        {
            if (index <= excelReader.ResultsCount - 1 && index >= 0)
            {
                int i = 0; while (i++ != index) excelReader.NextResult();
            }
        }
        public void goToSheet(int index)
        {
            resetConnection();
            selectSheet(index);
            //cloceExcelReader();
        }
        public string sheetName(int index)
        {
            return excelSheet(index).TableName;
        }

        public int filedsCountCurrentSheet()
        {
            return excelReader.FieldCount;
        }
        public void loadFieldSheetName(int indexSheets)
        {
            fieldsheet.Clear();
            goToSheet(indexSheets);
            
            if (excelReader.Read())
            {
                for (int i = 0; i < excelReader.FieldCount; i++)
                    fieldsheet.Add(excelReader.GetValue(i).ToString());
            }
        }

        public List<string> filedNameSheet()
        {
            return fieldsheet;
        }

        private bool read(DataTable table, ) 
        {
            return !excelReader.Read();
        }

        public bool readAllFile()
        {
            bool f;
            try
            {
                f = resetConnection();

                for (int i = 0; i < dataFile.Tables.Count; ++i)
                    dataFile.Tables[i].Clear();
                dataFile.Clear();

                dataFile = excelReader.AsDataSet(true);
            }
            catch (System.Exception Except)
            {
                MessageBox.Show(Except.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return f;
        }
        public bool readAtSheet(int indexSheet, int[] columnIndex)
        {
            string item = string.Empty;
            try
            {
                if (worksheet.Count == 0) loadSheetsName();

                if (worksheet.Count > 0 && indexSheet >= 0 && indexSheet <= worksheet.Count - 1)
                {
                    item = worksheet[indexSheet];
                    if (dataFile.Tables.Contains(item))
                    {
                        dataFile.Tables[item].Clear();
                        dataFile.Tables[item].Columns.Clear();
                    }
                    else
                        dataFile.Tables.Add(new DataTable(item));
                }
                else { return false; }

                DataTable table = dataFile.Tables[item];

                if (excelReader.IsFirstRowAsColumnNames == true)
                {
                    loadFieldSheetName(indexSheet);
                    for (int i = 0; i < columnIndex.Length; i++)
                    {
                        try
                        {
                            table.Columns.Add(fieldsheet[columnIndex[i]]);
                        }
                        catch (Exception)
                        {
                            table.Columns.Add(fieldsheet[columnIndex[i]] + "_" + i.ToString());
                        }
                    }
                }
                else
                {
                    goToSheet(indexSheet);

                    for (int i = 0; i < columnIndex.Length; i++)
                        table.Columns.Add("Column " + i.ToString());
                }

                while (excelReader.Read())
                {
                    DataRow row = table.NewRow();
                    for (int i = 0; i < columnIndex.Length; i++)
                        row[i] = excelReader.GetValue(columnIndex[i]);
                    table.Rows.Add(row);
                }
                table.AcceptChanges();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dataFile.Tables[item].Clear();
                return false;
            }
            return !excelReader.Read();
        }
        public bool readAtSheet(int indexSheet)
        {
            string item = string.Empty;
            try
            {
                if (worksheet.Count == 0) loadSheetsName();

                if (worksheet.Count > 0 && indexSheet >= 0 && indexSheet <= worksheet.Count - 1)
                {
                    item = worksheet[indexSheet];
                    if (dataFile.Tables.Contains(item))
                        dataFile.Tables[item].Clear();
                    else
                        dataFile.Tables.Add(new DataTable(item));
                }
                else { return false; }

                DataTable table = dataFile.Tables[item];
                table.Columns.Clear();
                loadFieldSheetName(indexSheet);
                goToSheet(indexSheet);

                if (excelReader.IsFirstRowAsColumnNames == true)
                {
                    for (int i = 0; i < fieldsheet.Count; i++)
                    {
                        try
                        {
                            table.Columns.Add(fieldsheet[i]);
                        }
                        catch (Exception)
                        {
                            table.Columns.Add(fieldsheet[i] + "_" + i.ToString());
                        }
                    }
                    excelReader.Read();
                }
                else
                {
                    for (int i = 0; i < fieldsheet.Count; i++)
                        table.Columns.Add("Column " + i.ToString());
                }
                
                while (excelReader.Read())
                {
                    DataRow row = table.NewRow();
                    for (int i = 0; i < fieldsheet.Count; i++)
                        row[i] = excelReader.GetValue(i);
                    table.Rows.Add(row);
                }
                table.AcceptChanges();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dataFile.Tables[item].Clear();
                return false;
            }
            return !excelReader.Read();
        }

        public DataTable excelSheet(string Name)
        {
            try
            {
                //int index = worksheet.IndexOf(Name);
                int index = dataFile.Tables.IndexOf(Name);
                if (index >= 0 && (index <= dataFile.Tables.Count - 1))
                    return dataFile.Tables[index];
                else throw new System.Exception("Произошло обращение к не существующему листу Excel файла");
            }
            catch (System.Exception Except)
            {
                MessageBox.Show(Except.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public DataTable excelSheet(int index)
        {
            try
            {
                string item = string.Empty;
                if ((index <= worksheet.Count - 1) && index >= 0)
                     item = worksheet[index];
                else throw new System.Exception("Произошло обращение к не существующему листу Excel файла");

                int trueIndex = dataFile.Tables.IndexOf(item);

                if ((trueIndex <= dataFile.Tables.Count - 1) && trueIndex >= 0)
                    return dataFile.Tables[trueIndex];
                else throw new System.Exception("Произошло обращение к не существующему листу Excel файла");
            }
            catch (System.Exception Except)
            {
                MessageBox.Show(Except.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public IExcelDataReader getReader()
        {
            return excelReader;
        }

        public void clearData()
        {
            dataFile.Clear();
            worksheet.Clear();
        }

    }
}
