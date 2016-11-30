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
    public partial class Form8 : Form
    {
        public Form8()
        {
            InitializeComponent();
        }

        private void Form8_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Form3 f3 = new Form3();
            f3.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string newpath = null;
            using (var dialog = new FolderBrowserDialog())
                if (dialog.ShowDialog() == DialogResult.OK)
                    newpath = dialog.SelectedPath;
                else
                {
                    MessageBox.Show("Новая директория не выбрана!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    newpath = AppDomain.CurrentDomain.BaseDirectory;
                    return;
                }
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
                                        childnode.InnerText = newpath;
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

        private void button3_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Form9 f9 = new Form9();
            f9.ShowDialog();
            this.Visible = true;
        }
    }
}
