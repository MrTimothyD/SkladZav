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
    public partial class Form5 : Form
    {
        public Form5()
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
        }

        private void Form5_Load(object sender, EventArgs e)
        {
            try
            {
                this.orderTableAdapter.Fill(this.skladZavDBDataSet3.Order);
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

        private void Form5_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.orderTableAdapter.Fill(this.skladZavDBDataSet3.Order);
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
