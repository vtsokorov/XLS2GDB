using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace xls2gdb
{
    public partial class MainForm : Form
    {
        private Database db = new Database();
        private ExcelReader reader = new ExcelReader();
        private IndexConteiner indexCont = new IndexConteiner();

        //Список таблиц
        private List<string> exTableList = new List<string>();
        private List<string> dbTableList = new List<string>();

        //Индекс текущей таблицы
        private int indexExcelTable    = -1;
        private int indexfirebirdTable = -1;
        
        //Список полей
        private List<string> dbFields = new List<string>();
        private List<string> exFields = new List<string>();

        private DataTable tableExport;
        private DataTable tableImport;
        int runIndex  =-1;

        private int indexToMoveExcelFileds;
        private int indexToMoveFirebirdFields;


//----------------------------------------------------------------------------
        private void exitbutton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        public MainForm()
        {
            InitializeComponent();
        }

        //user define function!
        private string[] splitString(string item)
        {
            char[] chars = new char[4];
            chars[0] = ' '; chars[1] = '-';
            chars[2] = '>'; chars[3] = ' ';
            return item.Split(chars, StringSplitOptions.RemoveEmptyEntries);
        }
        private int validationFields(CheckedListBox leftWidget, CheckedListBox rightWidget, out List<int> leftArray, out List<int> rightArray)
        {
            int leftCount = leftWidget.CheckedItems.Count;
            int rightCount = rightWidget.CheckedItems.Count;
            
            leftArray = new List<int>();
            rightArray = new List<int>();

            if ((leftCount != rightCount) || leftCount == 0 || rightCount == 0)
                return -1;

            for (int i = 0; i < leftCount; ++i)
            {
                string leftItem  = leftWidget.CheckedItems[i].ToString();
                string rightItem = rightWidget.CheckedItems[i].ToString();
                
                int leftIndex  = leftWidget.Items.IndexOf(leftItem);
                int rightIndex = rightWidget.Items.IndexOf(rightItem);

                if (leftIndex == rightIndex)
                {
                    int indexE = exFields.IndexOf(leftItem);
                    int indexF = dbFields.IndexOf(rightItem);

                    leftArray.Add(indexE);
                    rightArray.Add(indexF);
                }
                else
                {
                    leftArray.Clear(); rightArray.Clear();
                    return -2;
                }
            }
            return 0;
        }
        private Dictionary<int, int> validationFields(CheckedListBox leftWidget, CheckedListBox rightWidget)
        {
            int leftCount = leftWidget.CheckedItems.Count;
            int rightCount = rightWidget.CheckedItems.Count;

            Dictionary<int, int> dic = new Dictionary<int, int>();

            if ((leftCount != rightCount) || leftCount == 0 || rightCount == 0)
                return dic;

            for (int i = 0; i < leftCount; ++i)
            {
                string leftItem = leftWidget.CheckedItems[i].ToString();
                string rightItem = rightWidget.CheckedItems[i].ToString();

                int leftIndex = leftWidget.Items.IndexOf(leftItem);
                int rightIndex = rightWidget.Items.IndexOf(rightItem);

                if (leftIndex == rightIndex)
                {
                    int indexE = exFields.IndexOf(leftItem);
                    int indexF = dbFields.IndexOf(rightItem);
                    dic.Add(indexE, indexF);
                }
            }
            return dic;
        }
        private int getCheckedTable(CheckedListBox widget)
        {
            int index = -1;
            for (int i = 0; i < widget.Items.Count; ++i)
            {
                if (widget.GetItemChecked(i))
                {
                    index = i;
                    break;
                }
            }
            return index;
        }
        private void ItemCheckWidgwt(CheckedListBox widget, ItemCheckEventArgs e, Button button)
        {
            if (e.NewValue == CheckState.Checked)
            {
                for (int i = 0; i < widget.Items.Count; ++i)
                    if (e.Index != i)
                        widget.SetItemChecked(i, false);
                button.Enabled = true;
            }
            else button.Enabled = false;
        }

        private void upField(CheckedListBox widget)
        {
            if (widget.SelectedItems.Count > 0)
            {
                object selected = widget.SelectedItem;
                int indx = widget.Items.IndexOf(selected);
                int totl = widget.Items.Count;
                bool flag = widget.GetItemChecked(indx);
                if (indx == 0)
                {
                    widget.Items.Remove(selected);
                    widget.Items.Insert(totl - 1, selected);
                    widget.SetSelected(totl - 1, true);
                    widget.SetItemChecked(0, flag);
                }
                else
                {
                    widget.Items.Remove(selected);
                    widget.Items.Insert(indx - 1, selected);
                    widget.SetSelected(indx - 1, true);
                    widget.SetItemChecked(indx - 1, flag);
                }
            }
        }
        private void downField(CheckedListBox widget) 
        {
            if (widget.SelectedItems.Count > 0)
            {
                object selected = widget.SelectedItem;
                int indx = widget.Items.IndexOf(selected);
                int totl = widget.Items.Count;
                bool flag = widget.GetItemChecked(indx);
                if (indx == totl - 1)
                {
                    widget.Items.Remove(selected);
                    widget.Items.Insert(0, selected);
                    widget.SetSelected(0, true);
                    widget.SetItemChecked(0, flag);
                }
                else
                {
                    widget.Items.Remove(selected);
                    widget.Items.Insert(indx + 1, selected);
                    widget.SetSelected(indx + 1, true);
                    widget.SetItemChecked(indx + 1, flag);
                }
            }
        }
        private void moveItem(CheckedListBox widget, DragEventArgs e, int index)
        {
            //индекс, куда перемещаем
            //listBox1.PointToClient(new Point(e.X, e.Y)) - необходимо
            //использовать поскольку в e храниться
            //положение мыши в экранных коородинатах, а эта
            //функция позволяет преобразовать в клиентские
            if (index != -1)
            {
                int newIndex = widget.IndexFromPoint(widget.PointToClient(new Point(e.X, e.Y)));
                bool flag = widget.GetItemChecked(index);
                //если вставка происходит в начало списка
                if (newIndex == -1)
                {
                    //получаем перетаскиваемый элемент
                    object itemToMove = widget.Items[index];
                    //удаляем элемент
                    widget.Items.RemoveAt(index);
                    //добавляем в конец списка
                    widget.Items.Add(itemToMove);
                    widget.SetItemChecked(widget.Items.Count - 1, flag);
                }
                //вставляем где-то в середину списка
                else if (index != newIndex)
                {
                    //получаем перетаскиваемый элемент
                    object itemToMove = widget.Items[index];
                    //удаляем элемент
                    widget.Items.RemoveAt(index);
                    //вставляем в конкретную позицию
                    widget.Items.Insert(newIndex, itemToMove);
                    widget.SetItemChecked(newIndex, flag);
                }
            }
            
        }
        //----------------------------
        private void MainForm_Load(object sender, EventArgs e)
        {
            ToolTip toolTip = new ToolTip();
            toolTip.AutoPopDelay = 5000;
            toolTip.InitialDelay = 1000;
            toolTip.ReshowDelay = 500;
            toolTip.ShowAlways = true;
            toolTip.SetToolTip(excelFieldsList, "Переместите элемент мышью");
            toolTip.SetToolTip(firebirdFieldsList, "Переместите элемент мышью");
            portNumericUpDown.Value = 3050;

            showExcelTableButton.Enabled = false;
            showFirebirdTableButton.Enabled = false;
            prevButton.Enabled = false;

            serverTypeCombo.SelectedIndex = 0;
            charsetCombo.SelectedIndex    = 26;

            passwordText.Text = "masterkey";
            roleText.Text = "ADMINISTRATOR";
            fileClientLibraryPath.Text = @".\fbclient.dll";

            filedbPathText.Text = @"D:\workspace\db\ROV_DATABASE\BASE\AbitPodrazdel2012.gdb"; //"C:\Users\root\Documents\Visual Studio 2012\Projects\xls2gdb-Alpha-V3\xls2gdb-Alpha-V3\bin\Debug\XPDATA.GDB";
        }

        private void nextButton_Click(object sender, EventArgs e)
        {
            bool nextStep = false;
            switch (wizardControl.SelectedIndex+1)
            {
                case 0: { nextStep = firstStep(); break; }
                case 1: { nextStep = secondStep(); break; }
                case 2: { nextStep = threeStep();  break; }
                case 3: { nextStep = fourStep(); break; }
                case 4: { MessageBox.Show("5"); break; }
                default: { MessageBox.Show("-1");  break; }
            }

            if (nextStep && wizardControl.SelectedIndex + 1 <= wizardControl.TabPages.Count)
                wizardControl.SelectedIndex++;
        }
        private void prevButton_Click(object sender, EventArgs e)
        {
            bool nextStep = false;
            switch (wizardControl.SelectedIndex-1)
            {
                case 0: { nextStep = firstStep(); break; }
                case 1: { nextStep = true; break; }
                case 2: { nextStep = true; break; }
                case 3: { nextStep = fourStep(); break; }
                case 4: { MessageBox.Show("5"); break; }
                default: { MessageBox.Show("-1"); break; }
            }

            if (wizardControl.SelectedIndex - 1 >= 0)
                wizardControl.SelectedIndex--;
        }

        private void selectionChange(object sender, EventArgs e)
        {
            if (serverTypeCombo.SelectedIndex == 1)
            {
                serverNameLabel.Visible = true;
                serverNameCombo.Visible = true;
            }
            if (serverTypeCombo.SelectedIndex == 0)
            {
                serverNameLabel.Visible = false;
                serverNameCombo.Visible = false;
            }
        }

        private void selectExcelFileButton_Click(object sender, EventArgs e)
        {
            openFileDialog.InitialDirectory = @"c:\";
            openFileDialog.Filter = "книга Excel 97-2003 (*.xls)|*.xls|книга Excel (*.xlsx)|*.xlsx";
            openFileDialog.FilterIndex = 2;
            openFileDialog.FileName = "";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
                fileExcelPath.Text = openFileDialog.FileName;
        }
        private void selectFirebirdFileButton_Click(object sender, EventArgs e)
        {
            openFileDialog.InitialDirectory = @"c:\";
            openFileDialog.Filter = "firebird database (*.fdb)|*.fdb|InterBase database (*.gdb)|*.gdb";
            openFileDialog.FilterIndex = 2;
            openFileDialog.FileName = "";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
                filedbPathText.Text = openFileDialog.FileName;
        }
        private void selectClientLibraryFileButton_Click(object sender, EventArgs e)
        {
            openFileDialog.InitialDirectory = @".\";
            openFileDialog.Filter = "Client library (*.dll)|*.dll|Все файлы (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
            openFileDialog.FileName = "fbclient.dll";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
                fileClientLibraryPath.Text = openFileDialog.FileName;
        }

        private void serverNameSetItem(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '\r')
            {
                if(serverNameCombo.Items.IndexOf(serverNameCombo.Text) == -1)
                    serverNameCombo.Items.Add(serverNameCombo.Text);
            }
        }
        private void setUserNameItem(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '\r')
            {
                if (userCombo.Items.IndexOf(userCombo.Text) == -1)
                    userCombo.Items.Add(userCombo.Text);
            }
        }

        private void testConnectdbbutton_Click(object sender, EventArgs e)
        {
            try
            {
                string path_db = string.Empty;
                if (serverTypeCombo.SelectedIndex == 1) {
                    if (serverNameCombo.Text.Length > 0)
                        path_db = serverNameCombo.Text + ":" + filedbPathText.Text;
                    else {
                        MessageBox.Show("Не указано имя сервера", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                if (serverTypeCombo.SelectedIndex == 0)
                    path_db = filedbPathText.Text;

                if (db.testConnection(fileClientLibraryPath.Text, path_db, userCombo.Text, passwordText.Text, roleText.Text, charsetCombo.Text, (int)portNumericUpDown.Value))
                    MessageBox.Show("Соединение осуществлено успешно.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception Except)
            {
                MessageBox.Show(Except.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void testReadingExcelDataButton_Click(object sender, EventArgs e)
        {
            if (fileExcelPath.Text.Length > 0)
            {
                reader.setSourceData(fileExcelPath.Text, firstRowIsHeader.Checked);
                if (reader.openExcelReader())
                {
                    MessageBox.Show("Тест пройдено успешно.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                    MessageBox.Show("Не удается открыть файл.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
                MessageBox.Show("Укажите путь к файлу.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void excelTablesChanget(object sender, EventArgs e)
        {
            excelFieldsList.Items.Clear();
            indexExcelTable = excelTablesComboBox.SelectedIndex;

            reader.loadFieldsName(indexExcelTable);
            reader.cloceExcelReader();
            exFields = reader.filedsNameList(indexExcelTable);
            for (int i = 0; i < exFields.Count; ++i)
                excelFieldsList.Items.Add(exFields[i]);
            showExcelTableButton.Enabled = true;
        }
        private void firebirdTablesChanget(object sender, EventArgs e)
        {
            firebirdFieldsList.Items.Clear();
            indexfirebirdTable = firebirdTablesComboBox.SelectedIndex;
            string tableName = firebirdTablesComboBox.Items[indexfirebirdTable].ToString();

            db.loadTableFields(tableName);
            dbFields = db.fieldsNmaeList(indexfirebirdTable);
            DataTable table =  db.dataTable(tableName);
            for (int i = 0; i < dbFields.Count; ++i)
            {
                firebirdFieldsList.Items.Add(dbFields[i]);
                firebirdFieldsList.SetItemChecked(i, !table.Columns[i].AllowDBNull);
            }
            showFirebirdTableButton.Enabled = true;
        }

        private void showExcelTableButton_Click(object sender, EventArgs e)
        {
            reader.readAtSheet(indexExcelTable);
            string item = excelTablesComboBox.Items[indexExcelTable].ToString();
            DataTable table = reader.excelSheet(item);

            ExcelTable dialog = new ExcelTable();
            dialog.showTable(table);
            dialog.Show();
        }
        private void showFirebirdTableButton_Click(object sender, EventArgs e)
        {
            if (!db.isOpen()) db.openDataBase();

            string item = firebirdTablesComboBox.Items[indexfirebirdTable].ToString();
            string selectCommand = "SELECT * FROM " + item;
            db.select(db.initAdapter(item, selectCommand, false));
            DataTable table = db.dataTable(item);

            DataBaseTable dialog = new DataBaseTable();

            dialog.showTable(table);
            dialog.Show();
            db.closeDataBase();
        }

//-----Excel fileds DragDrop-----
        private void excelFieldsMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                indexToMoveExcelFileds = excelFieldsList.IndexFromPoint(e.X, e.Y);
                excelFieldsList.DoDragDrop(indexToMoveExcelFileds, DragDropEffects.Move);
            }
        }
        private void excelFieldsDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }
        private void excelFieldsDragDrop(object sender, DragEventArgs e)
        {
            moveItem(excelFieldsList, e, indexToMoveExcelFileds);
        }
//--firebird db fileds DragDrop--
        private void dbFieldsMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                indexToMoveFirebirdFields = firebirdFieldsList.IndexFromPoint(e.X, e.Y);
                firebirdFieldsList.DoDragDrop(indexToMoveFirebirdFields, DragDropEffects.Move);
            }
        }
        private void dbFieldsDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }
        private void dbFieldsDragDrop(object sender, DragEventArgs e)
        {
            moveItem(firebirdFieldsList, e, indexToMoveFirebirdFields);
        }
//------------------------------
        private void excelFieldUpButton_Click(object sender, EventArgs e)
        {
            upField(excelFieldsList);
        }
        private void excelFieldDownButton_Click(object sender, EventArgs e)
        {
            downField(excelFieldsList);
        }
        private void firebirdFieldUpButton_Click(object sender, EventArgs e)
        {
            upField(firebirdFieldsList);
        }
        private void firebirdFieldDownButton_Click(object sender, EventArgs e)
        {
            downField(firebirdFieldsList);
        }
//------------------------------
        private void saveIndexButton_Click(object sender, EventArgs e)
        {
            Dictionary<int, int> temp = validationFields(excelFieldsList, firebirdFieldsList);

            if (temp.Count == 0)
                MessageBox.Show("Индексы полей не сопоставлены.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                string leftTableName = excelTablesComboBox.Items[indexExcelTable].ToString();
                string rightTableName = firebirdTablesComboBox.Items[indexfirebirdTable].ToString();
                string item = leftTableName + " -> " + rightTableName;
                if (!transferComboBox.Items.Contains(item))
                {
                    transferComboBox.Items.Add(item);
                    transferComboBox.SelectedIndex = transferComboBox.Items.Count - 1;
                    runIndex = temp.Count;
                    indexCont.add(temp);
                }
                else
                {
                    int index = transferComboBox.Items.IndexOf(item);
                    transferComboBox.Items.RemoveAt(index);
                    transferComboBox.Items.Add(item);
                    transferComboBox.SelectedIndex = transferComboBox.Items.Count - 1;
                    indexCont.delete(index);
                    indexCont.add(temp);
                }
            }
        }
        private void deleteIndexButton_Click(object sender, EventArgs e)
        {
            if (transferComboBox.SelectedIndex >= 0)
            {
                indexCont.delete(transferComboBox.SelectedIndex);
                transferComboBox.Items.RemoveAt(transferComboBox.SelectedIndex);
                if (transferComboBox.Items.Count > 0)
                    transferComboBox.SelectedIndex = 0;
            }
        }

        private bool firstStep()
        {
            showExcelTableButton.Enabled = false;
            showFirebirdTableButton.Enabled = false;
            prevButton.Enabled = false;

            return true;
        }

        private bool secondStep()
        {
            excelFieldsList.Items.Clear();
            firebirdFieldsList.Items.Clear();

            if (fileExcelPath.Text.Length > 0)
            {
                reader.setSourceData(fileExcelPath.Text, firstRowIsHeader.Checked);
                reader.loadSheetsName();
                exTableList.Clear();
                excelTablesComboBox.Items.Clear();
                exTableList = reader.sheetsNameList();
                for (int i = 0; i < exTableList.Count; ++i)
                    excelTablesComboBox.Items.Add(exTableList[i]);
            }
            else {
                MessageBox.Show("Укажите путь к Excel файлу", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false; 
            }
            
            try
            {
                string path_db = string.Empty;
                if (filedbPathText.Text.Length == 0) 
                {
                    MessageBox.Show("Укажите путь к файлу базы данных", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                if (fileClientLibraryPath.Text.Length == 0)
                {
                    MessageBox.Show("Укажите путь к файлу клиентской библиотеки", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                if (serverTypeCombo.SelectedIndex == 1)
                {
                    if (serverNameCombo.Text.Length > 0)
                        path_db = serverNameCombo.Text + ":" + filedbPathText.Text;
                    else
                    {
                        MessageBox.Show("Не указано имя сервера", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
                if (serverTypeCombo.SelectedIndex == 0)
                    path_db = filedbPathText.Text;

                db.InitConnectString(fileClientLibraryPath.Text, path_db, userCombo.Text, passwordText.Text, roleText.Text, charsetCombo.Text, (int)portNumericUpDown.Value);
                if (db.openDataBase())
                {
                    dbTableList.Clear();
                    firebirdTablesComboBox.Items.Clear();
                    dbTableList = db.tablesNameList();
                    for (int i = 0; i < dbTableList.Count; ++i)
                        firebirdTablesComboBox.Items.Add(dbTableList[i]);
                }
                db.closeDataBase();
                prevButton.Enabled = true;
                return true;
            }
            catch (System.Exception Except)
            {
                MessageBox.Show(Except.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private bool threeStep()
        {
            if (runIndex == 0)
            {
                MessageBox.Show("Индексы полей не сопоставлены.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        private bool fourStep()
        {
            //int[] excelColumnIndex = excelIndexArray.ToArray();
            //int[] dbColumnIndex = firebirdIndexArray.ToArray();

            //reader.readAtSheet(indexExcelTable, excelColumnIndex);
            //string item = excelTablesComboBox.Items[indexExcelTable].ToString();
            //tableExport = reader.excelSheet(item);

            //List<string> fields = new List<string>();
            //for (int i = 0; i < dbColumnIndex.Length; ++i)
            //    fields.Add(dbFields[dbColumnIndex[i]]);

            //string tableName = firebirdTablesComboBox.Items[indexfirebirdTable].ToString();
            //string selectCommand = db.buildSimpleSelectCommand(tableName, fields.ToArray(), string.Empty);

            //if (!db.isOpen()) db.openDataBase();

            //db.select(db.initAdapter(tableName, selectCommand, true));
            //tableImport = db.dataTable(tableName);
            //db.closeDataBase();

            return true;
        }

        private bool fiveStep()
        {
            return true;
        }




    }
}
