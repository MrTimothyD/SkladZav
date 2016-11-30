using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Xml;

namespace SkladZav
{
    public partial class Form3 : Form
    {

        private Excel.Application excelapp;
        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;
        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;
        private Excel.Range excelcells;

        public Form3()
        {
            InitializeComponent();
                string path = "";
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(AppDomain.CurrentDomain.BaseDirectory + "SkladZav.xml");
                XmlElement xRoot = doc.DocumentElement;
                foreach (XmlNode xnode in xRoot)
                {
                    if (xnode.Attributes.Count > 0)
                    {
                        XmlNode attr = xnode.Attributes.GetNamedItem("name");
                        if (attr != null)
                            if (attr.Value.ToString() == "ConnectionString")
                            {
                                foreach (XmlNode childnode in xnode.ChildNodes)
                                {
                                    if (childnode.Name == "Path")
                                    {
                                        if (childnode.InnerText.ToString() != "")
                                            path = childnode.InnerText;
                                    }
                                }
                            }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось найти XML!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
            this.lastMonthTableAdapter.Connection.ConnectionString = path;
            this.halfRecordTableAdapter.Connection.ConnectionString = path;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Form4 f4 = new Form4();
            f4.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Form5 f5 = new Form5();
            f5.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(AppDomain.CurrentDomain.BaseDirectory + "SkladZav.xml");
                XmlElement xRoot = doc.DocumentElement;
                foreach (XmlNode xnode in xRoot)
                {
                    if (xnode.Attributes.Count > 0)
                    {
                        XmlNode attr = xnode.Attributes.GetNamedItem("name");
                        if (attr != null)
                            if (attr.Value.ToString() == "ExcelSaving")
                            {
                                foreach (XmlNode childnode in xnode.ChildNodes)
                                {
                                    if (childnode.Name == "Path")
                                    {
                                        if (childnode.InnerText.ToString() != "")
                                            path = childnode.InnerText;
                                    }
                                }
                            }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось найти XML!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
            try
            {
                this.halfRecordTableAdapter.Fill(this.skladZavDBDataSet4.HalfRecord);
                this.lastMonthTableAdapter.Fill(this.skladZavDBDataSet4.LastMonth);
            }
            catch (Exception)
            {
                this.Visible = false;
                MessageBox.Show("База данных не доступна! Нажмите OK и введите действительные настройки подключения.", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Error);             
                Form10 f10 = new Form10();
                f10.ShowDialog();
                return;
            }
            if (DateTime.Now.Month > Convert.ToDateTime(skladZavDBDataSet4.LastMonth.Rows[0]["Number"].ToString()).Month || ((Convert.ToDateTime(skladZavDBDataSet4.LastMonth.Rows[0]["Number"].ToString()).Month == 12) && (DateTime.Now.Month != Convert.ToDateTime(skladZavDBDataSet4.LastMonth.Rows[0]["Number"].ToString()).Month)))
            {
                int x = 2, i = 0, SumK = 2;
                decimal Sum = 0;
                this.Visible = false;
                try
                {
                    excelapp = new Excel.Application();
                    excelapp.Visible = false;
                    excelapp.SheetsInNewWorkbook = 1;
                    excelapp.Workbooks.Add(Type.Missing);
                    excelapp.DisplayAlerts = false;
                    excelappworkbooks = excelapp.Workbooks;
                    excelappworkbook = excelappworkbooks[1];
                    excelsheets = excelappworkbook.Worksheets;
                    excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

                    excelcells = (Excel.Range)excelworksheet.Columns["A:D", Type.Missing];
                    excelcells.ColumnWidth = 40;
                    excelcells.RowHeight = 23;
                    excelcells.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelcells.Cells.Font.Name = "Times New Roman";
                    excelcells.Cells.Font.Size = 12;
                    excelcells.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    excelcells.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                    excelcells = excelworksheet.get_Range("A1", "D1");
                    excelcells.RowHeight = 50;
                    excelcells.Cells.Font.Size = 14;
                    excelcells.Cells.Font.Bold = 1;
                    excelcells.Cells.Interior.ColorIndex = 6;

                    excelcells = excelworksheet.get_Range("A1", Type.Missing);
                    excelcells.Cells.Value = "Наименование";

                    excelcells = excelworksheet.get_Range("B1", Type.Missing);
                    excelcells.Cells.Value = "Количество по заявке";

                    excelcells = excelworksheet.get_Range("C1", Type.Missing);
                    excelcells.Cells.Value = "Количество по приходу";

                    excelcells = excelworksheet.get_Range("D1", Type.Missing);
                    excelcells.Cells.Value = "Сумма";

                    excelcells = (Excel.Range)excelworksheet.Columns["D:D", Type.Missing];
                    excelcells.Cells.NumberFormat = "#,##0.00$";

                    excelworksheet.get_Range("A" + x, "D" + x).Cells.Interior.ColorIndex = 4;
                    excelworksheet.get_Range("A" + x, Type.Missing).Cells.Value = "Заказы";
                    x = x + 1;
                    DataRow[] SubunitRow = this.skladZavDBDataSet4.Tables["halfrecord"].Select("Subunit = 'Заказ'");
                    while (i + 1 <= SubunitRow.Length)
                    {
                        excelworksheet.get_Range("A" + x, Type.Missing).Cells.Value = SubunitRow[i]["Name"].ToString();
                        excelworksheet.get_Range("B" + x, Type.Missing).Cells.Value = Convert.ToInt32(SubunitRow[i]["CountR"].ToString());
                        excelworksheet.get_Range("C" + x, Type.Missing).Cells.Value = Convert.ToInt32(SubunitRow[i]["CountIN"].ToString());
                        excelworksheet.get_Range("D" + x, Type.Missing).Cells.Value = Convert.ToDecimal(SubunitRow[i]["Sum"].ToString());
                        Sum = Sum + Convert.ToDecimal(SubunitRow[i]["Sum"].ToString());
                        x = x + 1;
                        i++;
                    }
                    if (Sum != 0)
                    {
                        excelworksheet.get_Range("D" + SumK, Type.Missing).Cells.Value = Sum;
                    }
                    i = 0;

                    excelworksheet.get_Range("A" + x, "D" + x).Cells.Interior.ColorIndex = 4;
                    excelworksheet.get_Range("A" + x, Type.Missing).Cells.Value = "Цех";
                    SumK = x;
                    Sum = 0;
                    x = x + 1;
                    SubunitRow = this.skladZavDBDataSet4.Tables["halfrecord"].Select("Subunit = 'Цех'");
                    while (i + 1 <= SubunitRow.Length)
                    {
                        excelworksheet.get_Range("A" + x, Type.Missing).Cells.Value = SubunitRow[i]["Name"].ToString();
                        excelworksheet.get_Range("B" + x, Type.Missing).Cells.Value = Convert.ToInt32(SubunitRow[i]["CountR"].ToString());
                        excelworksheet.get_Range("C" + x, Type.Missing).Cells.Value = Convert.ToInt32(SubunitRow[i]["CountIN"].ToString());
                        excelworksheet.get_Range("D" + x, Type.Missing).Cells.Value = Convert.ToDecimal(SubunitRow[i]["Sum"].ToString());
                        Sum = Sum + Convert.ToDecimal(SubunitRow[i]["Sum"].ToString());
                        x = x + 1;
                        i++;
                    }
                    if (Sum != 0)
                    {
                        excelworksheet.get_Range("D" + SumK, Type.Missing).Cells.Value = Sum;
                    }
                    i = 0;

                    excelworksheet.get_Range("A" + x, "D" + x).Cells.Interior.ColorIndex = 4;
                    excelworksheet.get_Range("A" + x, Type.Missing).Cells.Value = "ОГМ";
                    SumK = x;
                    Sum = 0;
                    x = x + 1;
                    SubunitRow = this.skladZavDBDataSet4.Tables["halfrecord"].Select("Subunit = 'ОГМ'");
                    while (i + 1 <= SubunitRow.Length)
                    {
                        excelworksheet.get_Range("A" + x, Type.Missing).Cells.Value = SubunitRow[i]["Name"].ToString();
                        excelworksheet.get_Range("B" + x, Type.Missing).Cells.Value = Convert.ToInt32(SubunitRow[i]["CountR"].ToString());
                        excelworksheet.get_Range("C" + x, Type.Missing).Cells.Value = Convert.ToInt32(SubunitRow[i]["CountIN"].ToString());
                        excelworksheet.get_Range("D" + x, Type.Missing).Cells.Value = Convert.ToDecimal(SubunitRow[i]["Sum"].ToString());
                        Sum = Sum + Convert.ToDecimal(SubunitRow[i]["Sum"].ToString());
                        x = x + 1;
                        i++;
                    }
                    if (Sum != 0)
                    {
                        excelworksheet.get_Range("D" + SumK, Type.Missing).Cells.Value = Sum;
                    }
                    i = 0;

                    excelworksheet.get_Range("A" + x, "D" + x).Cells.Interior.ColorIndex = 4;
                    excelworksheet.get_Range("A" + x, Type.Missing).Cells.Value = "Белоусов";
                    SumK = x;
                    Sum = 0;
                    x = x + 1;
                    SubunitRow = this.skladZavDBDataSet4.Tables["halfrecord"].Select("Subunit = 'Белоусов'");
                    while (i + 1 <= SubunitRow.Length)
                    {
                        excelworksheet.get_Range("A" + x, Type.Missing).Cells.Value = SubunitRow[i]["Name"].ToString();
                        excelworksheet.get_Range("B" + x, Type.Missing).Cells.Value = Convert.ToInt32(SubunitRow[i]["CountR"].ToString());
                        excelworksheet.get_Range("C" + x, Type.Missing).Cells.Value = Convert.ToInt32(SubunitRow[i]["CountIN"].ToString());
                        excelworksheet.get_Range("D" + x, Type.Missing).Cells.Value = Convert.ToDecimal(SubunitRow[i]["Sum"].ToString());
                        Sum = Sum + Convert.ToDecimal(SubunitRow[i]["Sum"].ToString());
                        x = x + 1;
                        i++;
                    }
                    if (Sum != 0)
                    {
                        excelworksheet.get_Range("D" + SumK, Type.Missing).Cells.Value = Sum;
                    }
                    i = 0;

                    excelworksheet.get_Range("A" + x, "D" + x).Cells.Interior.ColorIndex = 4;
                    excelworksheet.get_Range("A" + x, Type.Missing).Cells.Value = "Стройка";
                    SumK = x;
                    Sum = 0;
                    x = x + 1;
                    SubunitRow = this.skladZavDBDataSet4.Tables["halfrecord"].Select("Subunit = 'Стройка'");
                    while (i + 1 <= SubunitRow.Length)
                    {
                        excelworksheet.get_Range("A" + x, Type.Missing).Cells.Value = SubunitRow[i]["Name"].ToString();
                        excelworksheet.get_Range("B" + x, Type.Missing).Cells.Value = Convert.ToInt32(SubunitRow[i]["CountR"].ToString());
                        excelworksheet.get_Range("C" + x, Type.Missing).Cells.Value = Convert.ToInt32(SubunitRow[i]["CountIN"].ToString());
                        excelworksheet.get_Range("D" + x, Type.Missing).Cells.Value = Convert.ToDecimal(SubunitRow[i]["Sum"].ToString());
                        Sum = Sum + Convert.ToDecimal(SubunitRow[i]["Sum"].ToString());
                        x = x + 1;
                        i++;
                    }
                    if (Sum != 0)
                    {
                        excelworksheet.get_Range("D" + SumK, Type.Missing).Cells.Value = Sum;
                    }
                    i = 0;

                    excelworksheet.get_Range("A" + x, "D" + x).Cells.Interior.ColorIndex = 4;
                    excelworksheet.get_Range("A" + x, Type.Missing).Cells.Value = "Объекты";
                    SumK = x;
                    Sum = 0;
                    x = x + 1;
                    SubunitRow = this.skladZavDBDataSet4.Tables["halfrecord"].Select("Subunit = 'Объекты'");
                    while (i + 1 <= SubunitRow.Length)
                    {
                        excelworksheet.get_Range("A" + x, Type.Missing).Cells.Value = SubunitRow[i]["Name"].ToString();
                        excelworksheet.get_Range("B" + x, Type.Missing).Cells.Value = Convert.ToInt32(SubunitRow[i]["CountR"].ToString());
                        excelworksheet.get_Range("C" + x, Type.Missing).Cells.Value = Convert.ToInt32(SubunitRow[i]["CountIN"].ToString());
                        excelworksheet.get_Range("D" + x, Type.Missing).Cells.Value = Convert.ToDecimal(SubunitRow[i]["Sum"].ToString());
                        Sum = Sum + Convert.ToDecimal(SubunitRow[i]["Sum"].ToString());
                        x = x + 1;
                        i++;
                    }
                    if (Sum != 0)
                    {
                        excelworksheet.get_Range("D" + SumK, Type.Missing).Cells.Value = Sum;
                    }
                    i = 0;

                    excelworksheet.get_Range("A" + x, "D" + x).Cells.Interior.ColorIndex = 4;
                    excelworksheet.get_Range("A" + x, Type.Missing).Cells.Value = "Прочие";
                    SumK = x;
                    Sum = 0;
                    x = x + 1;
                    SubunitRow = this.skladZavDBDataSet4.Tables["halfrecord"].Select("Subunit = 'Прочие'");
                    while (i + 1 <= SubunitRow.Length)
                    {
                        excelworksheet.get_Range("A" + x, Type.Missing).Cells.Value = SubunitRow[i]["Name"].ToString();
                        excelworksheet.get_Range("B" + x, Type.Missing).Cells.Value = Convert.ToInt32(SubunitRow[i]["CountR"].ToString());
                        excelworksheet.get_Range("C" + x, Type.Missing).Cells.Value = Convert.ToInt32(SubunitRow[i]["CountIN"].ToString());
                        excelworksheet.get_Range("D" + x, Type.Missing).Cells.Value = Convert.ToDecimal(SubunitRow[i]["Sum"].ToString());
                        Sum = Sum + Convert.ToDecimal(SubunitRow[i]["Sum"].ToString());
                        x = x + 1;
                        i++;
                    }
                    if (Sum != 0)
                    {
                        excelworksheet.get_Range("D" + SumK, Type.Missing).Cells.Value = Sum;
                    }
                    i = 0;
                }
                catch (Exception)
                {
                    MessageBox.Show("Ошибка открытия Excel!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                try
                {
                    string get = path + @"\" + "Отчет с " + Convert.ToDateTime(skladZavDBDataSet4.LastMonth.Rows[0]["Number"].ToString()).ToShortDateString() + " по " + DateTime.Now.ToShortDateString() + ".xlsx";
                    excelappworkbook = excelappworkbooks[1];
                    excelappworkbook.SaveAs(get, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    excelapp.Workbooks.Close();
                    excelapp.Quit();
                    for (i = 0; i + 1 <= this.skladZavDBDataSet4.Tables["halfrecord"].Rows.Count; i++)
                    {
                        this.skladZavDBDataSet4.Tables["halfrecord"].Rows[i].Delete();
                    }
                    this.halfRecordTableAdapter.Update(this.skladZavDBDataSet4.HalfRecord);
                    skladZavDBDataSet4.LastMonth.Rows[0]["Number"] = DateTime.Now;
                    this.lastMonthTableAdapter.Update(this.skladZavDBDataSet4.LastMonth);
                }
                catch (Exception)
                {
                    excelapp.Workbooks.Close();
                    excelapp.Quit();
                    MessageBox.Show("Не удалось сохранить отчет!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                this.Visible = true;
            } 
        }

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Form7 f7 = new Form7();
            f7.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(AppDomain.CurrentDomain.BaseDirectory + "SkladZav.xml");
                XmlElement xRoot = doc.DocumentElement;
                foreach (XmlNode xnode in xRoot)
                {
                    if (xnode.Attributes.Count > 0)
                    {
                        XmlNode attr = xnode.Attributes.GetNamedItem("name");
                        if (attr != null)
                            if (attr.Value.ToString() == "ExcelSaving")
                            {
                                foreach (XmlNode childnode in xnode.ChildNodes)
                                {
                                    if (childnode.Name == "Path")
                                    {
                                        if (childnode.InnerText.ToString() != "")
                                        {
                                            path = childnode.InnerText;
                                        }
                                    }
                                }
                            }
                    }
                }
                System.Diagnostics.Process.Start("explorer", path);
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось найти XML!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }
    }
}
