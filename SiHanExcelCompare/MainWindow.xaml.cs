using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace SiHanExcelCompare
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private class DataValue
        {
            public IWorkbook workbook;
            public List<string> sheetNameList=new List<string>();
            public Dictionary<string, ISheet> sheetDic=new Dictionary<string, ISheet>();
            public List<SheetData> sheetList=new List<SheetData>();
            public List<HeaderData> headerList = new List<HeaderData>();
            public DataValue(IWorkbook workbook)
            {
                this.workbook = workbook;
                int j = 0;
                try
                {
                    while (workbook.GetSheetAt(j) != null)
                    {
                        string name = workbook.GetSheetName(j);
                        sheetNameList.Add(name);
                        sheetDic.Add(name, workbook.GetSheetAt(j));
                        sheetList.Add(new SheetData(j+1, name, 1, false));
                        if (j==0)
                        {
                            SelectSheet(sheetList[0]);
                        }
                        j++;
                    }
                }
                catch { }
            }
            public void SelectSheet(SheetData data)
            {
                for (int i = 0; i < sheetList.Count; i++)
                {
                    sheetList[i].selected = false;
                }
                data.selected = true;
                ResetLineNum();
            }
            public void ResetLineNum()
            {
                headerList.Clear();
                try
                {
                    var sheet = GetSelectSheet();
                    var sheetData = GetSelectSheetData();
                    var lineNum = sheetData.lineNum;
                    IRow firstRow = sheet.GetRow(lineNum-1);
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数
                    for (int j = 1, i = 0; i <= cellCount; i++, j++)
                    {
                        headerList.Add(new HeaderData(j, GetCellValue(firstRow, i)));
                    }
                }
                catch { }
            }
            public ISheet GetSelectSheet()
            {
                string name = "";
                for (int i = 0; i < sheetList.Count; i++)
                {
                    if (sheetList[i].selected)
                    {
                        name = sheetList[i].name;
                        break;
                    }
                }
                return sheetDic[name];
            }
            public int GetSelectSheetLastCoumnNum()
            {
                return headerList.Count;
            }
            public int GetSelectLineNum()
            {
                int lineNum = 0;
                for (int i = 0; i < sheetList.Count; i++)
                {
                    if (sheetList[i].selected)
                    {
                        lineNum = sheetList[i].lineNum-1;
                        break;
                    }
                }
                return lineNum;
            }
            public SheetData GetSelectSheetData()
            {
                for (int i = 0; i < sheetList.Count; i++)
                {
                    if (sheetList[i].selected)
                    {
                        return sheetList[i];
                    }
                }
                return null;
            }
        }
        class SheetData : INotifyPropertyChanged  
        {
            public int index;
            public string name;
            public int lineNum;
            public bool selected;
            public event PropertyChangedEventHandler PropertyChanged;
            public string Index
            {
                get
                {
                    return index.ToString();
                }
                set
                {
                    index = Convert.ToInt32(value);
                    if (this.PropertyChanged != null)//激发事件，参数为Age属性    
                    {
                        this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Index"));
                    }
                }
            }
            public string Name
            {
                get
                {
                    return name;
                }
                set
                {
                    name = value;
                    if (this.PropertyChanged != null)//激发事件，参数为Age属性    
                    {
                        this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Name"));
                    }
                }
            }
            public string LineNum
            {
                get
                {
                    return lineNum.ToString();
                }
                set
                {
                    lineNum = Convert.ToInt32(value);
                    if (this.PropertyChanged != null)//激发事件，参数为Age属性    
                    {
                        this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("LineNum"));
                    }
                }
            }
            public string Selected
            {
                get
                {
                    if (selected)
                    {
                        return "是";
                    }
                    else
                    {
                        return "";
                    }
                }
                set
                {
                    selected = Convert.ToBoolean(value);
                    if (this.PropertyChanged != null)//激发事件，参数为Age属性    
                    {
                        this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Selected"));
                    }
                }
            }
            public SheetData() { }
            public SheetData(int index, string name, int lineNum, bool selected)
            {
                this.index = index;
                this.name = name;
                this.lineNum = lineNum;
                this.selected = selected;
            }
        }
        class HeaderData : INotifyPropertyChanged
        {
            public int column;
            public string name;
            public int targetColumn;
            public event PropertyChangedEventHandler PropertyChanged;
            public string Column
            {
                get
                {
                    return column.ToString();
                }
                set
                {
                    column = Convert.ToInt32(value) ;
                    if (this.PropertyChanged != null)//激发事件，参数为Age属性    
                    {
                        this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Column"));
                    }
                }
            }
            public string Name
            {
                get
                {
                    return name;
                }
                set
                {
                    name =value;
                    if (this.PropertyChanged != null)//激发事件，参数为Age属性    
                    {
                        this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Name"));
                    }
                }
            }
            public string TargetColumn
            {
                get
                {
                    if (targetColumn<=0)
                    {
                        return "";
                    }
                    return targetColumn.ToString();
                }
                set
                {
                    targetColumn = Convert.ToInt32(value);
                    if (this.PropertyChanged != null)//激发事件，参数为Age属性    
                    {
                        this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("TargetClumn"));
                    }
                }
            }
            public HeaderData() { }
            public HeaderData(int column,string name)
            {
                this.column = column;
                this.name = name;
            }
        }

        private DataValue m_source_data;
        private DataValue m_target_data;
        private DataTable m_source_table=new DataTable();
        private DataTable m_source_table_compared = new DataTable();
        private DataTable m_target_table =new DataTable();
        private DataTable m_target_table_compared = new DataTable();
        private List<int> m_sourceEqList = new List<int>();
        private List<int> m_targetEqList = new List<int>();
        private bool m_showAll=true;
        public MainWindow()
        {
            InitializeComponent();
        }
        private void sourceBrowser_btn_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "文本文件|*.xlsx;*.xls";
            if (dialog.ShowDialog() == true)
            {
                ClearSourceData();
                FileInfo info = new FileInfo(dialog.FileName);
                sourceFileName_text.Text = info.Name.Replace(info.Extension,"");
                try
                {
                    using (var fs = new FileStream(info.FullName, FileMode.Open, FileAccess.Read))
                    {
                        if (info.FullName.IndexOf(".xlsx") > 0) // 2007版本
                            m_source_data = new DataValue(new XSSFWorkbook(fs));
                        else if (info.FullName.IndexOf(".xls") > 0) // 2003版本
                            m_source_data = new DataValue(new HSSFWorkbook(fs));
                    }
                    sourceSheetList.ItemsSource = null;
                    sourceHeaderList.ItemsSource = null;
                    sourceSheetList.ItemsSource = m_source_data.sheetList ;
                    sourceHeaderList.ItemsSource = m_source_data.headerList;
                    ResetSourceDataTable();
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.ToString(), "错误", MessageBoxButton.OK);
                }
            }
        }
        private void targetBrowser_btn_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "文本文件|*.xlsx;*.xls";
            if (dialog.ShowDialog() == true)
            {
                ClearTargetData();
                FileInfo info = new FileInfo(dialog.FileName);
                targetFileName_text.Text = info.Name.Replace(info.Extension, "");
                try
                {
                    using (var fs = new FileStream(info.FullName, FileMode.Open, FileAccess.Read))
                    {
                        if (info.FullName.IndexOf(".xlsx") > 0) // 2007版本
                            m_target_data = new DataValue(new XSSFWorkbook(fs));
                        else if (info.FullName.IndexOf(".xls") > 0) // 2003版本
                            m_target_data = new DataValue(new HSSFWorkbook(fs));
                    }
                    targetSheetList.ItemsSource = null;
                    targetHeaderList.ItemsSource = null;
                    targetSheetList.ItemsSource = m_target_data.sheetList;
                    targetHeaderList.ItemsSource = m_target_data.headerList;
                    ResetTargetDataTable();
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.ToString(), "错误", MessageBoxButton.OK);
                }
            }
        }

        private void sourceSheetList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //SheetData data = sourceSheetList.SelectedItem as SheetData;
            //if (data != null)
            //{
            //    m_source_data.SelectSheet(data);
            //    sourceHeaderList.ItemsSource = null;
            //    sourceHeaderList.ItemsSource = m_source_data.headerList;
            //}
        }
        private void sourceSheetList_ItemDoubleClick(object sender, MouseButtonEventArgs e)
        {
            SheetData data = sourceSheetList.SelectedItem as SheetData;
            if (data != null)
            {
                m_source_data.SelectSheet(data);
                sourceHeaderList.ItemsSource = null;
                sourceHeaderList.ItemsSource = m_source_data.headerList;
                sourceSheetList.ItemsSource = null;
                sourceSheetList.ItemsSource = m_source_data.sheetList;
                ResetSourceDataTable();
            }
        }
        private void sourceHeaderList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //HeaderData data = sourceHeaderList.SelectedItem as HeaderData;
            //if (data != null)
            //{

            //}
        }
        private void targetSheetList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //SheetData data = targetSheetList.SelectedItem as SheetData;
            //if (data != null)
            //{
            //    m_target_data.SelectSheet(data);
            //    targetHeaderList.ItemsSource = null;
            //    targetHeaderList.ItemsSource = m_target_data.headerList;
            //}
        }
        private void targetSheetList_ItemDoubleClick(object sender, MouseButtonEventArgs e)
        {
            SheetData data = targetSheetList.SelectedItem as SheetData;
            if (data != null)
            {
                m_target_data.SelectSheet(data);
                targetHeaderList.ItemsSource = null;
                targetHeaderList.ItemsSource = m_target_data.headerList;
                targetSheetList.ItemsSource = null;
                targetSheetList.ItemsSource = m_target_data.sheetList;
                ResetTargetDataTable();
            }
        }


        private void sourceSheetLineNum_text_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!isNumberic(e.Text))
            {
                e.Handled = true;
            }
            else
            {
                var s = Convert.ToInt32(e.Text);
                if (s < 1)
                {
                    e.Handled = true;
                }
                else
                {
                    e.Handled = false;
                }
            }
        }
        private void sourceSheetLineNum_text_TextChanged(object sender, TextChangedEventArgs e)
        {
        }
        private void targetSheetLineNum_text_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!isNumberic(e.Text))
            {
                e.Handled = true;
            }
            else
            {
                var s = Convert.ToInt32(e.Text);
                if (s < 1)
                {
                    e.Handled = true;
                }
                else
                {
                    e.Handled = false;
                }
            }
        }
        private void targetSheetLineNum_text_TextChanged(object sender, TextChangedEventArgs e)
        {
        }
        private void targetHeaderTargetColumn_text_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!isNumberic(e.Text))
            {
                e.Handled = true;
            }
            else
            {
                var s = Convert.ToInt32(e.Text);
                if (s < 0)
                {
                    e.Handled = true;
                }
                else
                    e.Handled = false;
            }
        }
        private bool isNumberic(string _string)
        {
            if (string.IsNullOrEmpty(_string))
                return false;
            foreach (char c in _string)
            {
                if (!char.IsDigit(c))
                    return false;
            }
            return true;
        }
        private static string GetCellValue(IRow row,int cellIndex)
        {
            try
            {
                var cell = row.GetCell(cellIndex);
                string columnValue = "";
                switch (cell.CellType)
                {
                    case CellType.Unknown:
                        break;
                    case CellType.Numeric:
                        columnValue = cell.NumericCellValue.ToString();
                        break;
                    case CellType.String:
                        columnValue = cell.StringCellValue;
                        break;
                    case CellType.Formula:
                        columnValue = cell.CellFormula;
                        break;
                    case CellType.Blank:
                        break;
                    case CellType.Boolean:
                        columnValue = cell.BooleanCellValue.ToString();
                        break;
                    case CellType.Error:
                        break;
                    default:
                        break;
                }
                return columnValue;
            }
            catch
            {
                return "";
            }
        }
        private static string GetCellValue(ISheet sheet,int rowIndex,int cellIndex)
        {
            try
            {
                var row = sheet.GetRow(rowIndex);
                return GetCellValue(row,cellIndex);
            }
            catch
            {
                return "";
            }
        }
        
        private void sourceTableData_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            DataGridRow dataGridRow = e.Row;
            var index = e.Row.GetIndex();
            var lineNum = m_source_data.GetSelectLineNum();
            if (m_showAll)
            {
                if (m_targetEqList.Count > 0 && !m_sourceEqList.Contains(index + lineNum + 1))
                {
                    dataGridRow.Background = Brushes.Plum;
                }
            }
            else
                dataGridRow.Background = Brushes.Plum;
        }
        private void targetTableData_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            DataGridRow dataGridRow = e.Row;
            var index = e.Row.GetIndex();
            var lineNum = m_target_data.GetSelectLineNum();
            if (m_showAll)
            {
                if (m_targetEqList.Count>0 && !m_targetEqList.Contains(index + lineNum + 1))
                {
                    dataGridRow.Background = Brushes.Plum;
                }
            }
            else
                dataGridRow.Background = Brushes.Plum;
        }
        private void Compare_btn_Click(object sender, RoutedEventArgs e)
        {
            if (m_source_data.headerList.Count <= 0 || m_target_data.headerList.Count <= 0)
            {
                MessageBox.Show("表格式不对，没有表单", "错误", MessageBoxButton.OK);
                return;
            }
            bool has = false;
            for (int i = 0; i < m_target_data.headerList.Count; i++)
            {
                if (m_target_data.headerList[i].targetColumn > 0)
                    has = true;
            }
            if (!has)
            {
                MessageBox.Show("要选择对比的表单编号", "错误", MessageBoxButton.OK);
                return;
            }
            m_source_table_compared = new DataTable();
            m_target_table_compared = new DataTable();
            SetComparedDataTable(m_source_data, m_target_data, ref m_source_table_compared, ref m_target_table_compared);
            int lastColumns = m_source_table.Columns.Count - 1;
            var lineNum = m_source_data.GetSelectLineNum();
            int cou = 0;
            for (int i = 0; i < m_source_table.Rows.Count; i++)
            {
                if (m_sourceEqList.Contains(i + lineNum + 1)) continue;
                var row = m_source_table.Rows[i];
                row.SetField(lastColumns, "--错误--");
                cou++;
            }
            sourceTableData.ItemsSource = null;
            targetTableData.ItemsSource = null;
            if (m_showAll)
            {
                sourceTableData.ItemsSource = m_source_table.DefaultView;
                targetTableData.ItemsSource = m_target_table.DefaultView;
            }
            else
            {
                sourceTableData.ItemsSource = m_source_table_compared.DefaultView;
                targetTableData.ItemsSource = m_target_table_compared.DefaultView;
            }
            var sourceSheet = m_source_data.GetSelectSheet();
            var targetSheet = m_target_data.GetSelectSheet();

            sourceResult_label.Content = "对比结果：  " + cou + "  个错误";
            //sourceTableData.Rows

        }
        private void SetDataTable(DataValue data, ref DataTable result, List<int> exludeList = null)
        {
            var sheet=data.GetSelectSheet();
            var sheetTop = data.GetSelectLineNum();
            SetTableColumns(sheet, sheetTop, ref result);
            for (int i = sheetTop+1; i <= sheet.LastRowNum; i++)
            {
                if (exludeList!=null && exludeList.Contains(i)) continue;
                var reRow=result.NewRow();
                try {
                    var row = sheet.GetRow(i);
                    for (int j = 0; j <= row.LastCellNum; j++)
                    {
                        var value = GetCellValue(sheet, i, j);
                        reRow.SetField<string>(j, value);
                    }
                }
                catch {
                }
                result.Rows.Add(reRow);
            }
        }
        private void SetComparedDataTable(DataValue source,DataValue target,ref DataTable sourceResult,ref DataTable targetResult)
        {
            var sourceSheet = source.GetSelectSheet();
            var sourceSheetTop = source.GetSelectLineNum();
            SetTableColumns(sourceSheet, sourceSheetTop, ref sourceResult);

            var targetSheet = target.GetSelectSheet();
            var targetSheetTop = target.GetSelectLineNum();
            SetTableColumns(targetSheet, targetSheetTop, ref targetResult);
            m_sourceEqList.Clear();
            m_targetEqList.Clear();
            for (int m = targetSheetTop+1; m <= targetSheet.LastRowNum; m++)
            {
                List<List<int>> waitCompareList = new List<List<int>>();
                for (int i = 0; i < target.headerList.Count; i++)
                {
                    var header = target.headerList[i];
                    if (header.targetColumn <= 0) continue;
                    var column = header.column - 1;
                    var needColumn = header.targetColumn - 1;
                    var value = GetCellValue(targetSheet, m, column);
                    waitCompareList.Add(GetSameRow(value, sourceSheet, needColumn));
                }
                if (waitCompareList.Count > 0)
                {
                    for (int i = 0; i < waitCompareList[0].Count; i++)
                    {
                        var wait = waitCompareList[0][i];
                        bool conta = true;
                        for (int j = 1; j < waitCompareList.Count; j++)
                        {
                            var w = waitCompareList[j];
                            if (!w.Contains(wait))
                            {
                                conta = false;
                            }
                        }
                        if (conta)
                        {
                            if (!m_sourceEqList.Contains(wait))
                            {
                                m_sourceEqList.Add(wait);
                            }
                            m_targetEqList.Add(m);
                        }
                    }
                }
            }
            SetDataTable(source, ref m_source_table_compared, m_sourceEqList);
            SetDataTable(target, ref m_target_table_compared, m_targetEqList);
        }
        private List<int> GetSameRow(string value,ISheet sheet,int cell)
        {
            var list = new List<int>();
            var length = sheet.LastRowNum;
            for (int i = 0; i <= length; i++)
            {
                try
                {
                    var row = sheet.GetRow(i);
                    if (value.Equals(GetCellValue(row,cell)))
                    {
                        list.Add(i);
                    }
                }
                catch
                {

                }
            }
            return list;
        }
        private void SetTableColumns(ISheet sheet,int rowIndex,ref DataTable result)
        {
            try
            {
                var row = sheet.GetRow(rowIndex);
                for (int i = 0; i <= row.LastCellNum; i++)
                {
                    var topValue = GetCellValue(row, i);
                    result.Columns.Add(topValue, typeof(string));
                }
            }
            catch
            {

            }
        }
        private void ResetSourceDataTable()
        {
            m_source_table = new DataTable();
            SetDataTable(m_source_data, ref m_source_table);
            sourceTableData.ItemsSource = null;
            sourceTableData.ItemsSource = m_source_table.DefaultView;
        }

        private void ResetTargetDataTable()
        {
            m_target_table = new DataTable();
            SetDataTable(m_target_data, ref m_target_table);
            targetTableData.ItemsSource = null;
            targetTableData.ItemsSource = m_target_table.DefaultView;
        }
        private void ClearSourceData()
        {
            m_source_data = null ;
            m_source_table = null;
            m_source_table_compared = new DataTable();
            m_sourceEqList.Clear();
            sourceTableData.ItemsSource = null;
            sourceSheetList.ItemsSource = null;
            sourceHeaderList.ItemsSource = null;
        }
        private void ClearTargetData()
        {
            m_target_data = null;
            m_target_table = null;
            m_target_table_compared = new DataTable();
            m_targetEqList.Clear();
            targetTableData.ItemsSource = null;
            targetSheetList.ItemsSource = null;
            targetHeaderList.ItemsSource = null;
        }

        private void showAll_ck_Unchecked(object sender, RoutedEventArgs e)
        {
            m_showAll = false;
            if (sourceTableData!=null)
            {
                sourceTableData.ItemsSource = null;
                targetTableData.ItemsSource = null;
                if (m_showAll)
                {
                    sourceTableData.ItemsSource = m_source_table.DefaultView;
                    targetTableData.ItemsSource = m_target_table.DefaultView;
                }
                else
                {
                    sourceTableData.ItemsSource = m_source_table_compared.DefaultView;
                    targetTableData.ItemsSource = m_target_table_compared.DefaultView;
                }
            }
        }

        private void showAll_ck_Checked(object sender, RoutedEventArgs e)
        {
            m_showAll = true;
            if (sourceTableData!=null)
            {
                sourceTableData.ItemsSource = null;
                targetTableData.ItemsSource = null;
                if (m_showAll)
                {
                    sourceTableData.ItemsSource = m_source_table.DefaultView;
                    targetTableData.ItemsSource = m_target_table.DefaultView;
                }
                else
                {
                    sourceTableData.ItemsSource = m_source_table_compared.DefaultView;
                    targetTableData.ItemsSource = m_target_table_compared.DefaultView;
                }
            }
        }
        private void export_btn_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog m_Dialog = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = m_Dialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            string dir = m_Dialog.SelectedPath.Trim();
            string filePath = dir +"/"+ sourceFileName_text.Text + "_结果.xlsx";
            int r=DataTableToExcel(m_source_table_compared, filePath, "sheet1", true);
            if (r>0)
                 MessageBox.Show("导出成功:"+filePath, "info", MessageBoxButton.OK);
        }


        private void export_btn_all_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog m_Dialog = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = m_Dialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            string dir = m_Dialog.SelectedPath.Trim();
            string filePath = dir + "/" + sourceFileName_text.Text + "_原.xlsx";
            int r = DataTableToExcel(m_source_table, filePath, "sheet1", true);
            if (r > 0)
                MessageBox.Show("导出成功:" + filePath, "info", MessageBoxButton.OK);
        }
        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcel(DataTable data,string filePath, string sheetName, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;
            IWorkbook workbook;
            if (filePath.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook();
            else if (filePath.IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook();
            else
                workbook = null;
            try
            {
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);
                }
                else
                {
                    return -1;
                }

                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
                using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    workbook.Write(fs); //写入到excel
                }
                return count;
            }
            catch (Exception ex)
            {
                MessageBox.Show("导出错误:"+ ex.Message, "错误", MessageBoxButton.OK);
                return -1;
            }
        }
    }
}
