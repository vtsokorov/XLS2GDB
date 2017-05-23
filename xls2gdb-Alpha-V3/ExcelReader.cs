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
        private bool loadSheets  = false;
        private FileStream stream;
        
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

        private void selectSheet(int index)
        {
            if (index <= excelReader.ResultsCount - 1 && index >= 0)
            {
                int i = 0; while (i++ != index) excelReader.NextResult();
            }
        }
        public void openSheet(int index)
        {
            resetConnection();
            selectSheet(index);
        }

        public int countSheets()
        {
            return excelReader.ResultsCount;
        }

        public int sheetIndex(string name)
        {
            return dataFile.Tables.IndexOf(name);
        }
        public void loadSheetsName()
        {
            resetConnection();
            dataFile.Tables.Clear();
            dataFile.Clear();

            for (int i = 0; i < excelReader.ResultsCount; ++i)
            {
                selectSheet(i);
                dataFile.Tables.Add(new DataTable(excelReader.Name));
            }

            loadSheets = true;

            cloceExcelReader();
        }
        public List<string> sheetsNameList()
        {
            List<string> list = new List<string>();
            for(int i = 0; i < dataFile.Tables.Count; ++i)
                list.Add(dataFile.Tables[i].TableName);
            return list;
        }
        public string sheetName(int index)
        {
            return excelSheet(index).TableName;
        }

        public DataTable excelSheet(string Name)
        {
            try
            {
                int index = dataFile.Tables.IndexOf(Name);
                if (index >= 0 && index <= dataFile.Tables.Count - 1)
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
                if (index >= 0 && index <= dataFile.Tables.Count - 1)
                    return dataFile.Tables[index];
                else throw new System.Exception("Произошло обращение к не существующему листу Excel файла");
            }
            catch (System.Exception Except)
            {
                MessageBox.Show(Except.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        
        public void loadFieldsName(int indexSheet)
        {
            openSheet(indexSheet);

            dataFile.Tables[indexSheet].Clear();
            dataFile.Tables[indexSheet].Columns.Clear();
            if (excelReader.Read())
            {
                if (excelReader.IsFirstRowAsColumnNames == true)
                {
                    for (int i = 0; i < excelReader.FieldCount; ++i)
                    {
                        string name = excelReader.GetValue(i).ToString();
                        if (dataFile.Tables[indexSheet].Columns.Contains(name))
                            dataFile.Tables[indexSheet].Columns.Add(name + "_" + i.ToString());
                        else
                            dataFile.Tables[indexSheet].Columns.Add(name);
                    }
                }
                else
                {
                    for (int i = 0; i < excelReader.FieldCount; i++)
                        dataFile.Tables[indexSheet].Columns.Add("Column_" + i.ToString());
                }
            }
            
            cloceExcelReader();
        }
        public void loadFieldsName(int indexSheet, int[] columnIndex)
        {
            openSheet(indexSheet);

            dataFile.Tables[indexSheet].Clear();
            dataFile.Tables[indexSheet].Columns.Clear();
            if (excelReader.Read())
            {
                if (excelReader.IsFirstRowAsColumnNames == true)
                {

                    for (int i = 0; i < columnIndex.Length; i++)
                    {
                        string name = excelReader.GetValue(columnIndex[i]).ToString();
                        if (dataFile.Tables[indexSheet].Columns.Contains(name))
                            dataFile.Tables[indexSheet].Columns.Add(name + "_" + i.ToString());
                        else
                            dataFile.Tables[indexSheet].Columns.Add(name);
                    }

                }
                else
                {
                    for (int i = 0; i < columnIndex.Length; i++)
                        dataFile.Tables[indexSheet].Columns.Add("Column_" + i.ToString());
                }
            }

            cloceExcelReader();
        }
        public List<string> filedsNameList(int indexSheet)
        {
            List<string> list = new List<string>();
            int count = dataFile.Tables[indexSheet].Columns.Count;

            for (int i = 0; i < count; ++i)
                list.Add(dataFile.Tables[indexSheet].Columns[i].ColumnName);
            return list;
        }

        private bool read(DataTable table) 
        {
            if (excelReader.IsFirstRowAsColumnNames == true)
                excelReader.Read();

            while (excelReader.Read())
            {
                DataRow row = table.NewRow();
                for (int i = 0; i < table.Columns.Count; i++)
                    row[i] = excelReader.GetValue(i);
                table.Rows.Add(row);
            }
            table.AcceptChanges();

            return !excelReader.Read();
        }
        private bool read(DataTable table, int[] columnIndex)
        {
            try
            {
                if (columnIndex.Length <= table.Columns.Count)
                {
                    if (excelReader.IsFirstRowAsColumnNames == true)
                        excelReader.Read();

                    while (excelReader.Read())
                    {
                        DataRow row = table.NewRow();
                        for (int i = 0; i < columnIndex.Length; i++)
                            row[i] = excelReader.GetValue(columnIndex[i]);
                        table.Rows.Add(row);
                    }
                    table.AcceptChanges();
                    return !excelReader.Read();
                }
                else throw new Exception("Размер массива индексов больше исходной таблицы");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public bool readAlltSheet()
        {
            try
            {
                if(resetConnection())
                {
                    dataFile.Clear();
                    dataFile = excelReader.AsDataSet(true);
                    loadSheets = true;
                    return !dataFile.HasErrors;
                }
                else 
                    return false;
            }
            catch (System.Exception Except)
            {
                MessageBox.Show(Except.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        public bool readAtSheet(int indexSheet)
        {
            try
            {
                if (!loadSheets) loadSheetsName();
                loadFieldsName(indexSheet);

                DataTable table = dataFile.Tables[indexSheet]; table.Clear();
               
                openSheet(indexSheet);

                bool f = read(table);
                return f;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        public bool readAtSheet(int indexSheet, int[] columnIndex)
        {
            try
            {
                if (!loadSheets) loadSheetsName();
                loadFieldsName(indexSheet, columnIndex);

                DataTable table = dataFile.Tables[indexSheet]; table.Clear();

                openSheet(indexSheet);

                bool f = read(table, columnIndex);
                return f;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
       
        public IExcelDataReader getReader()
        {
            return excelReader;
        }

        public void clearData()
        {
            dataFile.Clear();
        }

    }
}
