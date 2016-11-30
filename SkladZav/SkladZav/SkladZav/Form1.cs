using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace SkladZav
{
    public partial class Form1 : Form
    {
        public Form1()
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
            this.orderTableAdapter.Connection.ConnectionString = path;
            this.halfRecordTableAdapter.Connection.ConnectionString = path;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                this.orderTableAdapter.Fill(this.skladZavDBDataSet.Order);
            this.halfRecordTableAdapter.Fill(this.skladZavDBDataSet.HalfRecord);
            }
            catch (Exception)
            {
                MessageBox.Show("База данных не доступна! Нажмите OK и введите действительные настройки подключения.", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Visible = false;
                Form10 f10 = new Form10();
                f10.ShowDialog();
                return;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataGridView1.EndEdit();
                this.orderTableAdapter.Update(this.skladZavDBDataSet.Order);
            }
            catch (Exception)
            {
                MessageBox.Show("Не заполнены обязательные поля, либо введены неверные значения!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            int a;
            decimal b;
            if (Form2.Input(out a, out b))
                System.Windows.Forms.MessageBox.Show("Выполнено.", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
            {
                System.Windows.Forms.MessageBox.Show("Данные не были введены.", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                goto a1;
            }
            DataRow[] NameRow = this.skladZavDBDataSet.Tables["halfrecord"].Select("Name = '" + this.dataGridView1.CurrentRow.Cells[0].Value.ToString() + "'");
            DataRow[] SubunitRow = this.skladZavDBDataSet.Tables["halfrecord"].Select("Subunit = '" + this.dataGridView1.CurrentRow.Cells[1].Value.ToString() + "'");
            int i;
            for (i = 0; i < NameRow.Length; i++)
            {
                for (int j = 0; j < SubunitRow.Length; j++)
                {
                    if (NameRow[i] == SubunitRow[j])
                        goto update;
                }
            }
            // insert
            if (Convert.ToInt32(this.dataGridView1.CurrentRow.Cells[2].Value.ToString()) <= a)
            {
                this.halfRecordTableAdapter.Insert(this.dataGridView1.CurrentRow.Cells[0].Value.ToString(), this.dataGridView1.CurrentRow.Cells[1].Value.ToString(), a, Convert.ToInt32(this.dataGridView1.CurrentRow.Cells[2].Value.ToString()), a * b);
                this.dataGridView1.Rows.RemoveAt(this.dataGridView1.CurrentCell.RowIndex);
            }
            else
            {
                this.halfRecordTableAdapter.Insert(this.dataGridView1.CurrentRow.Cells[0].Value.ToString(), this.dataGridView1.CurrentRow.Cells[1].Value.ToString(), a, a, a * b);
                this.dataGridView1.CurrentRow.Cells[2].Value = Convert.ToInt32(this.dataGridView1.CurrentRow.Cells[2].Value.ToString()) - a;
            }
                goto a1;
        update:
            // update
            if (Convert.ToInt32(this.dataGridView1.CurrentRow.Cells[2].Value.ToString()) <= a)
            {
                NameRow[i]["CountIN"] = Convert.ToInt32(NameRow[i]["CountIN"].ToString()) + a;
                NameRow[i]["CountR"] = Convert.ToInt32(NameRow[i]["CountR"].ToString()) + Convert.ToInt32(this.dataGridView1.CurrentRow.Cells[2].Value.ToString());
                NameRow[i]["Sum"] = Convert.ToDecimal(NameRow[i]["Sum"].ToString()) + a * b;
                this.dataGridView1.Rows.RemoveAt(this.dataGridView1.CurrentCell.RowIndex);
            }
            else
            {
                NameRow[i]["CountIN"] = Convert.ToInt32(NameRow[i]["CountIN"].ToString()) + a;
                NameRow[i]["CountR"] = Convert.ToInt32(NameRow[i]["CountR"].ToString()) + a;
                NameRow[i]["Sum"] = Convert.ToDecimal(NameRow[i]["Sum"].ToString()) + a * b;
                this.dataGridView1.CurrentRow.Cells[2].Value = Convert.ToInt32(this.dataGridView1.CurrentRow.Cells[2].Value.ToString()) - a;
            }
        a1:
            this.halfRecordTableAdapter.Update(this.skladZavDBDataSet.HalfRecord);
            this.orderTableAdapter.Update(this.skladZavDBDataSet.Order);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.dataGridView1.EndEdit();
            this.orderTableAdapter.Update(this.skladZavDBDataSet.Order);
            Application.Exit();
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("Не заполнены обязательные поля, либо введены неверные значения!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.dataGridView1.Rows.RemoveAt(this.dataGridView1.CurrentCell.RowIndex);
            this.orderTableAdapter.Update(this.skladZavDBDataSet.Order);
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                if ((dataGridView1.CurrentRow.Cells[0].Value.ToString() != "") && (dataGridView1.CurrentRow.Cells[1].Value.ToString() != "") && (dataGridView1.CurrentRow.Cells[2].Value.ToString() != ""))
                {
                    
                    try
                    {
                        Validate();
                        this.dataGridView1.EndEdit();
                        this.orderTableAdapter.Update(this.skladZavDBDataSet.Order);
                        this.orderTableAdapter.Fill(this.skladZavDBDataSet.Order);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Не заполнены обязательные поля, либо введены неверные значения!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }
    }
}
