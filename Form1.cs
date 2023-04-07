using OfficeOpenXml;
using System.Data;
using System.Text;

namespace TestWinForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e) { }
        

        private void label2_Click(object sender, EventArgs e) { }
        
        private void button1_Click(object sender, EventArgs e)
        {
                        var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = @"EXCEL Files (*.xlsx)|*.xlsx|EXCEL Files 2003 (*.xls)|*.xls|All files (*.*)|*.*";
                  
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog.FileName;
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            if (maskedTextBox1.Text != "" && textBox1.Text != "" && textBox2.Text != "")
            {
                button1.Enabled = false;
                button2.Enabled = false;
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string OrderFilePath, WarehouseFilePath, FileNameResult;
                DataTable OrderDataTable, WarehouseDataTable, ResultDataTable = new DataTable();
                OrderFilePath = textBox1.Text;
                WarehouseFilePath = textBox2.Text;
                FileNameResult = "Result.xlsx";
                OrderDataTable = ReadXlsx(OrderFilePath);
                WarehouseDataTable = ReadXlsx(WarehouseFilePath);
                Dictionary<string?, int[]> dictionaryOrder = CreateDictionary(OrderDataTable);
                Dictionary<string?, int[]> dictionaryWarehouse = CreateDictionary(WarehouseDataTable);
                ResultDataTable = CompareTablesUpdate(WarehouseDataTable, OrderDataTable, ResultDataTable);
                ResultDataTable = FormatTable(ResultDataTable, OrderDataTable, WarehouseDataTable, OrderFilePath, WarehouseFilePath);
                WriteInFile(FileNameResult, ResultDataTable);
                
                GC.GetTotalMemory(true);
                Application.Exit();
            }
        }

        private void WriteInFile(string fileNameResult, DataTable resultDataTable)
        {
            FileInfo fileInfo = new FileInfo(fileNameResult);
            ExcelPackage p = new ExcelPackage(fileInfo);
            ExcelWorksheet myWorksheet = p.Workbook.Worksheets.First();
            myWorksheet.Cells[1, 1, 50000, 10].Value = null;
            for (int i = 1; i < resultDataTable.Rows.Count; i++)
            {

                for (int j = 1; j < resultDataTable.Columns.Count; j++)
                {
                    myWorksheet.Cells[i, j].Value = resultDataTable.Rows[i][j];
                }
            }

            p.Save();
        }

        private DataTable FormatTable(DataTable resultDataTable, DataTable orderDataTable, DataTable warehouseDataTable, string fileNameOrder, string fileNameWarehouse)
        {
            resultDataTable.Rows[3][1] = "N";
            for (int i = 4; i < resultDataTable.Rows.Count; i++)
            {
                resultDataTable.Rows[i][1] = i - 3;
            }


            resultDataTable.Rows[2][2] = fileNameOrder;
            resultDataTable.Rows[2][3] = fileNameOrder;
            resultDataTable.Rows[2][4] = fileNameWarehouse;
            resultDataTable.Rows[2][5] = fileNameWarehouse;

            StringBuilder sb = new StringBuilder();
            sb.Append(orderDataTable.Rows[0][0].ToString());
            sb.Append(" T1");
            resultDataTable.Rows[3][2] = sb;
            sb = new StringBuilder();
            sb.Append(orderDataTable.Rows[0][1].ToString());
            sb.Append(" T1");
            resultDataTable.Rows[3][3] = sb;
            sb = new StringBuilder();
            sb.Append(warehouseDataTable.Rows[0][0].ToString());
            sb.Append(" T2");
            resultDataTable.Rows[3][4] = sb;
            sb = new StringBuilder();
            sb.Append(warehouseDataTable.Rows[0][1].ToString());
            sb.Append(" T2");
            resultDataTable.Rows[3][5] = sb;

            resultDataTable.Rows[4][2] = 1;
            resultDataTable.Rows[4][3] = 2;
            resultDataTable.Rows[4][4] = 3;
            resultDataTable.Rows[4][5] = 4;


            for (var k = 6; k < resultDataTable.Columns.Count; k = k + 2)
            {
                resultDataTable.Rows[2][k] = fileNameOrder;
                resultDataTable.Rows[3][k] = orderDataTable.Rows[0][(k / 2) - 1].ToString();
                resultDataTable.Rows[4][k] = k;
            }
            for (var k = 7; k < resultDataTable.Columns.Count; k = k + 2)
            {
                resultDataTable.Rows[2][k] = fileNameWarehouse;
                resultDataTable.Rows[3][k] = warehouseDataTable.Rows[0][((k - 1) / 2) - 1].ToString();
                resultDataTable.Rows[4][k] = k;
            }
            return resultDataTable;

        }

        private DataTable CompareTablesUpdate(DataTable warehouseDataTable, DataTable orderDataTable, DataTable resultDataTable)
        {
            Dictionary<string?, int[]> dictionaryOrder = CreateDictionary(orderDataTable);
            Dictionary<string?, int[]> dictionaryWarehouse = CreateDictionary(warehouseDataTable);
            resultDataTable = CreateResultDataTable(warehouseDataTable, orderDataTable);
            DataRow row = resultDataTable.NewRow();
            resultDataTable.Rows.Add(row);
            row = resultDataTable.NewRow();
            resultDataTable.Rows.Add(row);
            row = resultDataTable.NewRow();
            resultDataTable.Rows.Add(row);
            row = resultDataTable.NewRow();
            resultDataTable.Rows.Add(row);
            row = resultDataTable.NewRow();
            resultDataTable.Rows.Add(row);

            for (int i = 2; i < warehouseDataTable.Rows.Count; i++)
            {
                int[] cell;
                dictionaryOrder.TryGetValue(warehouseDataTable.Rows[i][1].ToString(), out cell);

                if (cell != null)
                {

                    row = resultDataTable.NewRow();
                    resultDataTable.Rows.Add(row);
                    resultDataTable.Rows[resultDataTable.Rows.Count - 1][3] = orderDataTable.Rows[cell[0]][cell[1]].ToString();
                    resultDataTable.Rows[resultDataTable.Rows.Count - 1][2] = orderDataTable.Rows[cell[0]][cell[1] - 1].ToString();
                    resultDataTable.Rows[resultDataTable.Rows.Count - 1][5] = warehouseDataTable.Rows[i][1].ToString();
                    resultDataTable.Rows[resultDataTable.Rows.Count - 1][4] = warehouseDataTable.Rows[i][0].ToString();
                    for (int k = 6; k < resultDataTable.Columns.Count; k = k + 2)
                    {
                        resultDataTable.Rows[resultDataTable.Rows.Count - 1][k] = orderDataTable.Rows[cell[0]][(k / 2) - 1].ToString();
                    }
                    for (int k = 7; k < resultDataTable.Columns.Count; k = k + 2)
                    {
                        resultDataTable.Rows[resultDataTable.Rows.Count - 1][k] = warehouseDataTable.Rows[i][((k - 1) / 2) - 1].ToString();
                    }
                }
                else
                {
                    row = resultDataTable.NewRow();
                    resultDataTable.Rows.Add(row);
                    resultDataTable.Rows[resultDataTable.Rows.Count - 1][3] = null;
                    resultDataTable.Rows[resultDataTable.Rows.Count - 1][5] = warehouseDataTable.Rows[i][1].ToString();
                    resultDataTable.Rows[resultDataTable.Rows.Count - 1][4] = warehouseDataTable.Rows[i][0].ToString();
                    for (int k = 7; k < resultDataTable.Columns.Count; k = k + 2)
                    {
                        resultDataTable.Rows[resultDataTable.Rows.Count - 1][k] = warehouseDataTable.Rows[i][((k - 1) / 2) - 1].ToString();
                    }
                }


            }

            for (int i = 2; i < orderDataTable.Rows.Count; i++)
            {
                if (dictionaryWarehouse.ContainsKey(orderDataTable.Rows[i][1].ToString()))
                {
                }
                else
                {
                    row = resultDataTable.NewRow();
                    resultDataTable.Rows.Add(row);
                    resultDataTable.Rows[resultDataTable.Rows.Count - 1][5] = null;
                    resultDataTable.Rows[resultDataTable.Rows.Count - 1][3] = orderDataTable.Rows[i][1].ToString();
                    resultDataTable.Rows[resultDataTable.Rows.Count - 1][2] = orderDataTable.Rows[i][0].ToString();
                    for (int k = 6; k < resultDataTable.Columns.Count; k = k + 2)
                    {
                        resultDataTable.Rows[resultDataTable.Rows.Count - 1][k] = orderDataTable.Rows[i][(k / 2) - 1].ToString();
                    }
                }

            }

            return resultDataTable;
        }

        private DataTable CreateResultDataTable(DataTable warehouseDataTable, DataTable orderDataTable)
        {
            DataTable result = new DataTable();
            for (int i = 0; i < (warehouseDataTable.Columns.Count + orderDataTable.Columns.Count) + 2; i++)
            {
                DataColumn column = new DataColumn("Column " + i.ToString());
                result.Columns.Add(column);
            }
            return result;
        }

        private DataTable ReadXlsx(string filePath)
        {
            var pck = new OfficeOpenXml.ExcelPackage();
            pck.Load(File.OpenRead(filePath));
            var ws = pck.Workbook.Worksheets.First();
            DataTable resultDT = new DataTable();

            for (int i = 1; i < ws.Columns.Count(); i++)
            {
                if (ws.Cells[3, i].Value != null)
                {
                    DataColumn column = new DataColumn("Column " + i.ToString());
                    resultDT.Columns.Add(column);
                }
            }

            
            int rowCount = 50000;


            for (int i = 3; i < rowCount; i++)
            {

                if (ws.Cells[i, 1].Value != null)
                {
                    DataRow row = resultDT.NewRow();
                    //Console.WriteLine(ws.Cells[i, 1].Value + " " + resultDT.Rows.Count);

                    var wsRow = ws.Cells[i, 1, i, resultDT.Columns.Count];

                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }

                    resultDT.Rows.Add(row);
                }
                else
                {
                    rowCount = 0;
                }
            }

            pck.Dispose();
            return resultDT;
        }

        private Dictionary<string?, int[]> CreateDictionary(DataTable dt)
        {
            Dictionary<string?, int[]> newDictionary = new Dictionary<string?, int[]>();
            for (int i = 2; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][1] != dt.Rows[i - 1][1])
                    newDictionary.Add(dt.Rows[i][1].ToString(), new int[] { i, 1 });
            }
            return newDictionary;
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            string fileNameWarehouse;
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = @"EXCEL Files (*.xlsx)|*.xlsx|EXCEL Files 2003 (*.xls)|*.xls|All files (*.*)|*.*";

            if (openFileDialog.ShowDialog() != DialogResult.OK) return;
            textBox2.Text = openFileDialog.FileName;
            fileNameWarehouse = openFileDialog.FileName;
        }








        //private void JoinExcelFiles()
        //{
        //        string dataInd = DateTime.Now.ToString("dd.MM.yyyy HH mm ss");//���� ����� ������� ���������
        //        string pathExe = Application.StartupPath.ToString() + "\\";//���� � ����� exe
        //        string nameFolder = "";

        //        Microsoft.Office.Interop.Excel.Application excel;
        //        Microsoft.Office.Interop.Excel.Workbook wbInputExcel, wbResultExcel;//����� Excel
        //        Microsoft.Office.Interop.Excel.Worksheet wshInputExcel, wshResultExcel;//����� Excel
        //        //������� ����� � ������� ����� ��������� ��� ������������� ����� � ���������
        //        nameFolder = "��������_" + dataInd + "\\";

        //        if (Directory.Exists(pathExe + nameFolder))
        //        {
        //        }
        //        else
        //        {
        //            DirectoryInfo di = Directory.CreateDirectory(pathExe + nameFolder);
        //        }
        //        excel = new Microsoft.Office.Interop.Excel.Application();

        //        //�������� �������� ����� Excel, � ������� ����� ������������ ��� �����
        //        wbResultExcel = excel.Workbooks.Add(System.Reflection.Missing.Value);
        //        wshResultExcel = (Microsoft.Office.Interop.Excel.Worksheet)wbResultExcel.Sheets[1];
        //        wshResultExcel.Name = "����1";
        //        string nameFile = "Result.xlsx";
        //        wbResultExcel.SaveAs(pathExe + nameFolder + nameFile);

        //        excel.DisplayAlerts = false;

        //        string[] ourfiles = Directory.GetFiles(textBox1.Text, "*xls*", SearchOption.TopDirectoryOnly);

        //        //��������� ������ ���� � ��������� ��� ��������.
        //        wbInputExcel = excel.Workbooks.Open(ourfiles[0]);
        //        wshInputExcel = (Microsoft.Office.Interop.Excel.Worksheet)wbInputExcel.Worksheets.get_Item(1);
        //        int numInputRow = wshInputExcel.Rows.Count;
        //        //int numInputRow = wshInputExcel.Cells[wshInputExcel.Rows.Count, "A"].End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Row;

        //        wshInputExcel.Range[wshInputExcel.Cells[1, 1], wshInputExcel.Cells[numInputRow, Convert.ToInt32(maskedTextBox1.Text)]].Copy();

        //        //��������� � ������� ����
        //        wshResultExcel.Range["A1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
        //        wbResultExcel.Save();

        //        wbInputExcel.Close();

        //        ReleaseObject(wshInputExcel);
        //        ReleaseObject(wbInputExcel);

        //        int strBeginInput = 5; // "������ ��� ������ ������. ������ ������ � �������� �������, ��������� ������ � ���������� ������ �������, ������ �������� � 5"


        //        //���������� �� ��������� ������, �������� ������ � ��������� � ������� ����
        //        for (int f = 1; f <= ourfiles.Length - 1; f++)
        //        {
        //            wbInputExcel = excel.Workbooks.Open(ourfiles[f]);
        //            wshInputExcel = (Microsoft.Office.Interop.Excel.Worksheet)wbInputExcel.Worksheets.get_Item(1);


        //            //numInputRow = wshInputExcel.Cells[wshInputExcel.Rows.Count, "A"].End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Row;

        //            wshInputExcel.Range[wshInputExcel.Cells[strBeginInput, 1], wshInputExcel.Cells[numInputRow, Convert.ToInt32(maskedTextBox1.Text)]].Copy();

        //            //int numResultRow = wshResultExcel.Cells[wshResultExcel.Rows.Count, "A"].End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Row;
        //            int numResultRow = wshInputExcel.Rows.Count;
        //            wshResultExcel.Range[wshResultExcel.Cells[numResultRow + 1, 1], wshResultExcel.Cells[numResultRow + numInputRow, Convert.ToInt32(maskedTextBox1.Text)]].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
        //            wbInputExcel.Close(false);
        //        }
        //        //��������� �������� ����
        //        wbResultExcel.Save();
        //        wbResultExcel.Close(false);

        //        //����������� ����� � �������� ������
        //        MessageBox.Show("������");
        //        System.Diagnostics.Process.Start("explorer.exe", @"/select, " + pathExe + nameFolder);
        //    }

        //    private void ReleaseObject(object obj)
        //    {
        //        try
        //        {
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        //            obj = null;
        //        }
        //        catch (Exception ex)
        //        {
        //            obj = null;
        //            Console.WriteLine("Unable to release the Object {0}", ex.ToString());
        //        }
        //        finally
        //        {
        //            GC.Collect();
        //        }
        //    }
        //}
    }
}
