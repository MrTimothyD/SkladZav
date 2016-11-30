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
using System.Data.SqlClient;

namespace SkladZav
{
    public partial class Form9 : Form
    {
        public Form9()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text == "") || (textBox2.Text == "") || (textBox3.Text == ""))
            {
                MessageBox.Show("Заполните пустые поля!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                if (textBox2.Text != textBox3.Text)
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    MessageBox.Show("Пароли не совпадают!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    int? prov = 0;
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
                    try
                    {
                        var sqlConn = new SqlConnection(path);
                        {
                            var sqlCmd = new SqlCommand("UpdatePassword2", sqlConn);
                            sqlCmd.CommandType = CommandType.StoredProcedure;
                            sqlCmd.Parameters.AddWithValue("@old", textBox1.Text);
                            sqlCmd.Parameters.AddWithValue("@new", textBox2.Text);
                            sqlCmd.Parameters.Add("@prov", SqlDbType.Int, 4);
                            sqlCmd.Parameters["@prov"].Direction = ParameterDirection.Output;
                            sqlConn.Open();
                            sqlCmd.ExecuteScalar();
                            prov = Convert.ToInt32(sqlCmd.Parameters["@prov"].Value.ToString());
                            sqlConn.Close();
                        }
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("База данных не доступна! Нажмите OK и введите действительные настройки подключения.", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.Visible = false;
                        Form10 f10 = new Form10();
                        f10.ShowDialog();
                        return;
                    }
                    if (prov == 1)
                    {
                        textBox1.Text = null;
                        textBox2.Text = null;
                        textBox3.Text = null;
                        MessageBox.Show("Пароль изменен!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.Visible = false;
                    }
                    else
                    {
                        if (prov == 2)
                        {
                            textBox1.Text = null;
                            textBox2.Text = null;
                            textBox3.Text = null;
                            MessageBox.Show("Неверный пароль!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else
                        {
                            textBox1.Text = null;
                            textBox2.Text = null;
                            textBox3.Text = null;
                            MessageBox.Show("Пароль не был изменен!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            this.Visible = false;
        }

        private void Form9_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
            }
        }
    }
}
