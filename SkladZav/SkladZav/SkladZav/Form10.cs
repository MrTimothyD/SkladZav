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
    public partial class Form10 : Form
    {
        public Form10()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form10_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text == "") || (textBox2.Text == "") || (textBox3.Text == ""))
            {
                MessageBox.Show("Заполните пустые поля!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
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
                                            childnode.InnerText = "Data Source=" + textBox1.Text + ";Initial Catalog=" + textBox2.Text + ";Integrated Security=" + textBox3.Text;
                                            MessageBox.Show("Выполнено.", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        }
                                    }
                                }
                        }
                    }
                    doc.Save(AppDomain.CurrentDomain.BaseDirectory + "SkladZav.xml");
                }
                catch (Exception)
                {
                    MessageBox.Show("Не удалось найти XML!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();
                }
            }
            Application.Exit();
        }
    }
}
