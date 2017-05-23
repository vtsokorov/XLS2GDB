using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using FirebirdSql.Data.FirebirdClient;
using System.Text.RegularExpressions;

namespace xls2gdb
{
    class Database
    {
        
        private List<FbDataAdapter> adapters;
        private FbConnection connect;
        private DataSet localData;
        private int currentIndexTable;

        public FbConnection ConnectionObject
        {
            get { return connect; }
            set { connect = value; }
        }

        public Database()
        {
            connect   = new FbConnection();
            adapters  = new List<FbDataAdapter>();
            localData = new DataSet();
            currentIndexTable = -1;
        }
        //-------------------------------------------

        private string queryGetTables()
        {
            return "SELECT RDB$RELATION_NAME FROM RDB$RELATIONS WHERE (RDB$SYSTEM_FLAG = 0) AND (RDB$VIEW_SOURCE IS NULL) ORDER BY RDB$RELATION_NAME";
        }
        private string queryGetFields(string tableName)
        {
            return "SELECT R.RDB$FIELD_NAME, R.RDB$NULL_FLAG, T.RDB$TYPE_NAME, F.RDB$FIELD_LENGTH, F.RDB$FIELD_SCALE, F.RDB$FIELD_SUB_TYPE " +
                   "FROM RDB$TYPES T, RDB$RELATION_FIELDS R " +
                   "INNER JOIN RDB$FIELDS F ON F.RDB$FIELD_TYPE = T.RDB$TYPE AND T.RDB$FIELD_NAME = 'RDB$FIELD_TYPE' " +
                   "WHERE F.RDB$FIELD_NAME = R.RDB$FIELD_SOURCE AND R.RDB$SYSTEM_FLAG = 0 AND RDB$RELATION_NAME = '" + tableName +
                   "' ORDER BY R.RDB$FIELD_POSITION ASC";
        }
        private string queryGetTriggersForTable(string tableName)
        {
            return "SELECT RDB$TRIGGER_NAME FROM RDB$TRIGGERS WHERE RDB$RELATION_NAME = " + "'" + tableName + "'";
        }
        private string queryGetFirebidTypes()
        { 
            return "SELECT T.RDB$TYPE_NAME, T.RDB$TYPE FROM RDB$TYPES T WHERE T.RDB$FIELD_NAME = 'RDB$FIELD_TYPE' ORDER BY T.RDB$TYPE ASC";
        }
        private string queryGetFirebidType(string tableName, string fieldName)
        {
            return "SELECT FIRST 1 " + fieldName + " FROM " + tableName;
        }
        private DataColumn createDataColumn(string fieldName, string typeName, int subType, bool allowDBNull)
        {
            DataColumn column = new DataColumn(fieldName);
            column.AllowDBNull = allowDBNull;

            switch (typeName)
            {
                case ("SHORT")    : { column.DataType = typeof(Int16); break; }
                case ("LONG")     : { column.DataType = typeof(Int32); break; }
                case ("INT64")    : { column.DataType = subType == 0 ? typeof(Int64) : typeof(Decimal); break; }
                case ("FLOAT")    : { column.DataType = typeof(Single); break; }
                case ("DOUBLE")   : { column.DataType = typeof(Double); break; }
                case ("DATE")     : { column.DataType = typeof(DateTime); column.DateTimeMode = DataSetDateTime.Local; break; }
                case ("TIMESTAMP"): { column.DataType = typeof(DateTime); column.DateTimeMode = DataSetDateTime.Local; break; }
                case ("TIME")     : { column.DataType = typeof(TimeSpan); break; } 
                case ("TEXT")     : { column.DataType = typeof(String); break; }
                case ("VARYING")  : { column.DataType = typeof(String); break; }
                case ("CSTRING")  : { column.DataType = typeof(String); break; }
                case ("BLOB")     : { column.DataType = typeof(byte[]); break; }
            }
            return column;
        }

        public bool testConnection(string clientLibrary, string database, string userID, string password, string role, string charset, int port)
        {
            FbConnectionStringBuilder connectString = new FbConnectionStringBuilder();
            connectString.ClientLibrary = clientLibrary;
            connectString.Database = database;
            connectString.UserID = userID;
            connectString.Password = password;
            connectString.Charset = charset;
            connectString.Role = role;
            connectString.Port = port;
            connect.ConnectionString = connectString.ConnectionString;
            bool flag = openDataBase();
            connect.ConnectionString = string.Empty;
            if (flag == true) 
                closeDataBase();
            return flag;
        }
        public string InitConnectString(string clientLibrary, string database, string userID, string password, string role, string charset, int port)
        {
            FbConnectionStringBuilder connectString = new FbConnectionStringBuilder();
            connectString.ClientLibrary = clientLibrary;
            connectString.Database = database;
            connectString.UserID = userID;
            connectString.Password = password;
            connectString.Charset = charset;
            connectString.Role = role;
            connectString.Port = port;
            connect.ConnectionString = connectString.ConnectionString;
            connect.InfoMessage += new FbInfoMessageEventHandler(OnInfoMessage);
            this.loadTablelist();
            return connect.ConnectionString;
        }
  
        public bool openDataBase()
        {
            try
            {
                if (connect.ConnectionString.Length > 0)
                    connect.Open();
                else
                    throw new System.Exception("Строка соединения с БД не корректна.");
            }
            catch (System.Exception Except)
            {
                MessageBox.Show(Except.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connect.Close();
                return false;
            }
            return connect.State == System.Data.ConnectionState.Open ? true : false;
        }
        public bool closeDataBase()
        {
            try { connect.Close(); }
            catch (System.Exception Except)
            {
                MessageBox.Show(Except.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return connect.State == System.Data.ConnectionState.Closed ? true : false;
        }
        public bool isOpen()
        {
            return connect.State == System.Data.ConnectionState.Open ? true : false;
        }

        public void deactivateTriggers(string tableName, bool flag)
        {
            if (!isOpen()) openDataBase();

            FbCommand command = new FbCommand();
            command.Connection = connect;

            DataTable table = new DataTable("TRIGGERSNAME");
            FbDataAdapter adapter = new FbDataAdapter(queryGetTriggersForTable(tableName), connect);          
            adapter.Fill(table);

            string op = flag == true ? " INACTIVE" : " ACTIVE";

            int tableCount = table.Rows.Count;
            for (int i = 0, j = 0; i < tableCount; ++i)
            {
               string triggerName = table.Rows[i][j].ToString();
               if (triggerName != string.Empty)
               {
                   triggerName = triggerName.TrimEnd();
                   command.CommandText = "ALTER TRIGGER " + triggerName + op;
                   try {  command.ExecuteNonQuery(); }
                   catch (FbException e) {
                       MessageBox.Show("Неудалось деактевировать/активировать триггер: " + triggerName + "\n" + e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                   }
               }
            }

            if (isOpen()) closeDataBase();
        }

        //------------------------------------------
        public List<string> tablesNameList()
        {
            List<string> list = new List<string>();
            for (int i = 0; i < localData.Tables.Count; ++i)
                list.Add(localData.Tables[i].TableName);
            return list;
        }
        public List<string> fieldsNmaeList(int indexTable)
        {
            List<string> list = new List<string>();
            if (currentIndexTable == indexTable)
            {
                int count = localData.Tables[indexTable].Columns.Count;
                for (int i = 0; i < count; ++i)
                    list.Add(localData.Tables[indexTable].Columns[i].ColumnName);
            }
            return list;
        }

        private void loadTablelist()
        {
            if (!isOpen()) openDataBase();

            localData.Tables.Clear();
            localData.Clear();

            adapters.Clear();

            Regex rgx = new Regex(@"\$");
            DataTable table = new DataTable("TABLESNAME");
            FbDataAdapter adapter = new FbDataAdapter(queryGetTables(), connect);
            adapter.Fill(table);

            int tableCount = table.Rows.Count;
            for (int i = 0, j = 0; i < tableCount; ++i)
            {
                string item = table.Rows[i][j].ToString();
                item = item.TrimEnd();
                if (!rgx.IsMatch(item))
                {
                    adapters.Add(new FbDataAdapter());
                    localData.Tables.Add(new DataTable(item));
                }
            }
            if (isOpen()) closeDataBase();
        }  
        public void loadTableFields(string tableName)
        {
            int index = localData.Tables.IndexOf(tableName);
            if (index >= 0)
            {
                currentIndexTable = index;
                if (!isOpen()) openDataBase();

                DataTable table = new DataTable("FIELDSNAME");   
                FbDataAdapter adapter = new FbDataAdapter(queryGetFields(tableName), connect);
                adapter.Fill(table);
                int tableCount = table.Rows.Count;
                localData.Tables[index].Clear();
                localData.Tables[index].Columns.Clear();
                for (int i = 0; i < tableCount; ++i)
                {
                    string columnName = table.Rows[i][0].ToString().TrimEnd();
                    bool allowDBNull = table.Rows[i][1] != DBNull.Value ? false : true;
                    string typeName  = table.Rows[i][2].ToString().TrimEnd();
                    int subTypeIndex = table.Rows[i][5] != DBNull.Value ? Convert.ToInt32(table.Rows[i][5]) : 0;
                    DataColumn column = createDataColumn(columnName, typeName, subTypeIndex, allowDBNull);
                    localData.Tables[index].Columns.Add(column);
                }

                if (isOpen()) closeDataBase();
            }
        }

        public string initAdapter(string name, string commandText, bool clearColumn)
        {
            try
            {
                int index = localData.Tables.IndexOf(name);
                if (index >= 0 && commandText.Length > 0 && connect.ConnectionString.Length > 0)
                {
                    if (clearColumn == true)
                    {
                        localData.Tables[index].Clear();
                        localData.Tables[index].Columns.Clear();
                    }

                    adapters[index].SelectCommand = new FbCommand(commandText, connect);
                    return name;
                }
                else
                    throw new Exception("Ошибка! Проверте следующие параметры: имя таблици, строка соединения, команда выборки данных");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return string.Empty;
            }

        }

        public int tableIndex(string name)
        {
            try
            {
                int index = localData.Tables.IndexOf(name);
                if (index == -1)
                    throw new Exception("Таблицы с данным именем не существует");
                return index;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }
        }
        public string tableName(int index)
        {
            try
            {
                if (index >= 0 && index <= localData.Tables.Count - 1)
                    return localData.Tables[index].TableName;
                else
                    throw new Exception("Таблицы с данным индексом не существует");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return string.Empty;
            }
        }

        public DataTable dataTable(string tableName)
        {
            try
            {
                int index = localData.Tables.IndexOf(tableName);
                if (index >= 0 && index <= localData.Tables.Count - 1)
                    return localData.Tables[index];
                else
                    throw new Exception("Таблицы с данным именем не существует");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public FbDataAdapter tableAdapter(string tableName)
        {
            try
            {
                int index = localData.Tables.IndexOf(tableName);
                if (index >= 0 && index <= localData.Tables.Count - 1)
                    return adapters[index];
                else
                    throw new Exception("Адаптер таблицы с указанным именем не существует");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public bool refresh(string tableName)
        {
            try
            {
                int index = localData.Tables.IndexOf(tableName);
                if (index >= 0 && index <= localData.Tables.Count - 1)
                {
                    if (localData.Tables.Contains(tableName))
                        localData.Tables[index].Clear();
                    adapters[index].Fill(localData, tableName);
                    return true;
                }
                else
                    return false;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        public bool select(String tableName)
        {
            try
            {
                int index = localData.Tables.IndexOf(tableName);
                if (index >= 0 && index <= localData.Tables.Count - 1)
                {
                    if (localData.Tables.Contains(tableName))
                        localData.Tables[index].Clear();
                    adapters[index].Fill(localData, tableName);
                    return true;
                }
                else return false;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        public bool save(string tableName)
        {
            DataTable currentTable = dataTable(tableName);
            if (currentTable == null) return false;

            try
            {   
                tableAdapter(tableName).Update(currentTable.GetChanges());
                currentTable.AcceptChanges();
                return true;
            }
            catch (Exception)
            {
                currentTable.RejectChanges();
                return false;
            }
        }

        public void clear()
        {
            adapters.Clear();
            localData.Tables.Clear();
            localData.Clear();
            connect.InfoMessage -= new FbInfoMessageEventHandler(OnInfoMessage);
            connect.ConnectionString = "";
        }
        public void clear(int index)
        {
            try
            {
                string item = tableName(index);
                if (item == string.Empty)
                {
                    adapters.Remove(adapters[index]);
                    localData.Tables[index].Clear();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void clear(string tableName)
        {
            try
            {
                int index = tableIndex(tableName);
                if (index != -1)
                {
                    adapters.Remove(adapters[index]);
                    localData.Tables[index].Clear();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //-----------------------------------------

        public static void OnInfoMessage(object sender, FbInfoMessageEventArgs args)
        {
            List<FbError> errors = args.Errors.ToList();
            foreach (FbError err in errors)
            {
                StringBuilder str = new StringBuilder();
                str.AppendFormat("Ошибка серьезности {0}, ", err.Class)
                   .AppendFormat("номер ошибки {0} ", err.Number)
                   .AppendFormat("в строке {0} ", err.LineNumber)
                   .AppendFormat("{0}", err.Message);
                MessageBox.Show(str.ToString(), "Ощибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //-----------------------------------------
        public string buildSimpleSelectCommand(string tableName, string[] columns, string where)
        {
            StringBuilder selectCommand = new StringBuilder("SELECT ");
            if (columns.Length == 0)
                selectCommand.AppendFormat(" {0} ", "*");
            else for(int i = 0; i < columns.Length; ++i)
                selectCommand.AppendFormat(i != (columns.Length - 1) ? "\"{0}\", " : "\"{0}\" ", columns[i]);
            selectCommand.AppendFormat("{0} ", "FROM");
            selectCommand.AppendFormat(" \"{0}\"", tableName);

            selectCommand.AppendFormat((where.Length == 0) ? "{0}" : " WHERE {0}", where);

            return selectCommand.ToString();
        }
        public string buildSimpleInsertCommand(string tableName, string[] columns)
        {
            StringBuilder insertCommand = new StringBuilder("INSERT INTO");
            insertCommand.AppendFormat(" \"{0}\" (", tableName);
            for (int i = 0; i < columns.Length; ++i)
                insertCommand.AppendFormat(i != (columns.Length - 1) ? "\"{0}\", " : "\"{0}\"", columns[i]);
            insertCommand.AppendFormat("{0}", ")");
            insertCommand.AppendFormat(" {0}", "VALUES");
            insertCommand.AppendFormat(" {0}", "(");
            for (int i = 0; i < columns.Length; ++i)
                insertCommand.AppendFormat(i != (columns.Length - 1) ? "{0}, " : "{0}", "?");
            insertCommand.AppendFormat("{0}", ")");
            return insertCommand.ToString();
        }
        public string buildSimpleUpdateCommand(string tableName, string[] columns, string where)
        {
            StringBuilder updateCommand = new StringBuilder("UPDATE");
            updateCommand.AppendFormat(" \"{0}\" SET ", tableName);
            for (int i = 0; i < columns.Length; ++i)
                updateCommand.AppendFormat(i != (columns.Length - 1) ? "\"{0}\" = {1}, " : "\"{0}\" = {1}", columns[i], "?");
            updateCommand.AppendFormat((where.Length > 0) ? " WHERE {0}" : "{0}", where, "");
            return updateCommand.ToString();
        }
        public string buildSimpleDeleteCommand(string tableName, string[] columns, string where)
        {
            return string.Empty;
        }

        public FbParameter createFbParameter(DataColumn column)
        { 
            int size = 0; byte scale = 0;
            FbDbType fieldType;

            switch (column.DataType.ToString())
            {
                /*Array, Numeric, Text, Date*/
                 case ("System.Boolean") : { fieldType = FbDbType.Boolean;   break; }                
                 case ("System.Int16")   : { fieldType = FbDbType.SmallInt;  break; }
                 case ("System.Int32")   : { fieldType = FbDbType.Integer;   break; }
                 case ("System.Int64")   : { fieldType = FbDbType.BigInt;    break; }
                 case ("System.Single")  : { fieldType = FbDbType.Float;     break; }
                 case ("System.Double")  : { fieldType = FbDbType.Double;    break; }
                 case ("System.DateTime"): { fieldType = FbDbType.TimeStamp; break; }
                 case ("System.TimeSpan"): { fieldType = FbDbType.Time;      break; }
                 case ("System.byte[]")  : { fieldType = FbDbType.Binary;    break; }
                 case ("System.Guid")    : { fieldType = FbDbType.Guid;      break; }
                 case ("System.Char")    : { fieldType = FbDbType.Char;    size = column.MaxLength; break; }
                 case ("System.String")  : { fieldType = FbDbType.VarChar; size = column.MaxLength; break; }
                 case ("System.Decimal") : { fieldType = FbDbType.Decimal; size = 15; scale = 2; break; }
                default : { fieldType = FbDbType.VarChar; size = column.MaxLength; break; }
            }
            
            FbParameter params_db = new FbParameter(column.ColumnName, fieldType, size, column.ColumnName);
            params_db.Scale = scale;
            return params_db;
        }
        public void createInsertParameters(string tableName, params string[] columns)
        {
            try
            {
                FbDataAdapter adapter = tableAdapter(tableName);
                if (adapter == null)
                    throw new Exception("Адаптер даннной таблици не существует");

                if (adapter.InsertCommand == null) {
                    adapter.InsertCommand = new FbCommand();
                    adapter.InsertCommand.Connection = connect;
                }

                for (int i = 0; i < columns.Length; ++i)
                {
                    DataColumn column = dataTable(tableName).Columns[columns[i]];
                    FbParameter params_db   = createFbParameter(column);
                    params_db.Direction     = ParameterDirection.Input;
                    params_db.SourceVersion = DataRowVersion.Current;
                    adapter.InsertCommand.Parameters.Add(params_db);  
                }

                adapter.InsertCommand.CommandText = buildSimpleInsertCommand(tableName, columns);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void createUpdateParemeters(string tableName, string where, params string[] columns)
        {
            try
            {
                FbDataAdapter adapter = tableAdapter(tableName);
                if (adapter == null)
                    throw new Exception("Адаптер даннной таблици не существует");

                if (adapter.UpdateCommand == null) {
                    adapter.UpdateCommand = new FbCommand();
                    adapter.UpdateCommand.Connection = connect;
                }

                for (int i = 0; i < columns.Length; ++i)
                {
                    DataColumn column = dataTable(tableName).Columns[columns[i]];

                    FbParameter params_db = createFbParameter(column);
                    params_db.SourceVersion = DataRowVersion.Current;
                    adapter.UpdateCommand.Parameters.Add(params_db);

                    //FbParameter params_db = createFbParameter(column);
                    //params_db.SourceVersion = DataRowVersion.Original;
                    //adapter.UpdateCommand.Parameters.Add(params_db);
                }
                adapter.UpdateCommand.CommandText = buildSimpleUpdateCommand(tableName, columns, where);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void createDeleteParemeters(string tableName, string where, params string[] columns)
        {

        }
    }
}

/*
 -Firebird-	      -System-			   -ADO.NET-	
-----------------------------------------------------------------
SMALLINT		    SHORT  			    Int16
INTEGER			    LONG			    Int32
BIGINT			    INT64			    Int64
FLOAT           	FLOAT   		    Single
DOUBLE PRECISION  	DOUBLE 			    Double
NUMERIC          	INT64			    Decimal		
DECIMAL			    INT64			    Decimal		
DATE    		    DATE 			    DataTime	
TIME      	        TIME  			    TimeSpan
TIMESTAMP 	        TIMESTAMP		    DataTime	
CHAR			    TEXT    		    String	
VARCHAR			    VARYING (CSTRING)	String	
BLOB 			    BLOB 	 		    byte[]
------------------------------------------------------------------
QUAD                  
BLOB_ID      
*/