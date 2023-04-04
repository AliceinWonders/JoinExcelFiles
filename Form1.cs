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
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = fbd.SelectedPath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (maskedTextBox1.Text != "" && textBox1.Text != "")
            {
                button1.Enabled = false;
                button2.Enabled = false;
                //������� ����������� ������
                JoinExcelFiles();
                GC.GetTotalMemory(true);
                Application.Exit();
            }
        }

        private void JoinExcelFiles()
        {
            string dataInd = DateTime.Now.ToString("dd.MM.yyyy HH mm ss");//���� ����� ������� ���������
            string pathExe = Application.StartupPath.ToString() + "\\";//���� � ����� exe
            string nameFolder = "";

            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook wbInputExcel, wbResultExcel;//����� Excel
            Microsoft.Office.Interop.Excel.Worksheet wshInputExcel, wshResultExcel;//����� Excel
            //������� ����� � ������� ����� ��������� ��� ������������� ����� � ���������
            nameFolder = "��������_" + dataInd + "\\";

            if (Directory.Exists(pathExe + nameFolder))
            {
            }
            else
            {
                DirectoryInfo di = Directory.CreateDirectory(pathExe + nameFolder);
            }
            excel = new Microsoft.Office.Interop.Excel.Application();

            //�������� �������� ����� Excel, � ������� ����� ������������ ��� �����
            wbResultExcel = excel.Workbooks.Add(System.Reflection.Missing.Value);
            wshResultExcel = (Microsoft.Office.Interop.Excel.Worksheet)wbResultExcel.Sheets[1];
            wshResultExcel.Name = "����1";
            string nameFile = "Result.xlsx";
            wbResultExcel.SaveAs(pathExe + nameFolder + nameFile);

            excel.DisplayAlerts = false;

            string[] ourfiles = Directory.GetFiles(textBox1.Text, "*xls*", SearchOption.TopDirectoryOnly);
            
            //��������� ������ ���� � ��������� ��� ��������.
            wbInputExcel = excel.Workbooks.Open(ourfiles[0]);
            wshInputExcel = (Microsoft.Office.Interop.Excel.Worksheet)wbInputExcel.Worksheets.get_Item(1);
            int numInputRow = wshInputExcel.Rows.Count;
            //int numInputRow = wshInputExcel.Cells[wshInputExcel.Rows.Count, "A"].End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Row;

            wshInputExcel.Range[wshInputExcel.Cells[1, 1], wshInputExcel.Cells[numInputRow, Convert.ToInt32(maskedTextBox1.Text)]].Copy();

            //��������� � ������� ����
            wshResultExcel.Range["A1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
            wbResultExcel.Save();

            wbInputExcel.Close();

            ReleaseObject(wshInputExcel);
            ReleaseObject(wbInputExcel);

            int strBeginInput = 5; // "������ ��� ������ ������. ������ ������ � �������� �������, ��������� ������ � ���������� ������ �������, ������ �������� � 5"


            //���������� �� ��������� ������, �������� ������ � ��������� � ������� ����
            for (int f = 1; f <= ourfiles.Length - 1; f++)
            {
                wbInputExcel = excel.Workbooks.Open(ourfiles[f]);
                wshInputExcel = (Microsoft.Office.Interop.Excel.Worksheet)wbInputExcel.Worksheets.get_Item(1);

                
                //numInputRow = wshInputExcel.Cells[wshInputExcel.Rows.Count, "A"].End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Row;

                wshInputExcel.Range[wshInputExcel.Cells[strBeginInput, 1], wshInputExcel.Cells[numInputRow, Convert.ToInt32(maskedTextBox1.Text)]].Copy();

                //int numResultRow = wshResultExcel.Cells[wshResultExcel.Rows.Count, "A"].End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Row;
                int numResultRow = wshInputExcel.Rows.Count;
                wshResultExcel.Range[wshResultExcel.Cells[numResultRow + 1, 1], wshResultExcel.Cells[numResultRow + numInputRow, Convert.ToInt32(maskedTextBox1.Text)]].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
                wbInputExcel.Close(false);
            }
            //��������� �������� ����
            wbResultExcel.Save();
            wbResultExcel.Close(false);
                    
            //����������� ����� � �������� ������
            MessageBox.Show("������");
            System.Diagnostics.Process.Start("explorer.exe", @"/select, " + pathExe + nameFolder);
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Unable to release the Object {0}", ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}