﻿using System;
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
    public partial class Form4 : Form
    {
        public Form4()
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
            this.passwordTableAdapter.Connection.ConnectionString = path;
        }

        private void label2_Click(object sender, EventArgs e)
        {
            textBox1.Text = null;
            this.Visible = false;
            Form6 f6 = new Form6();
            f6.ShowDialog();
            this.passwordTableAdapter.Fill(this.skladZavDBDataSet1.Password);
            this.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = null;
            this.Visible = false;
            Form3 f3 = new Form3();
            f3.Visible = true;            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.passwordTableAdapter.Fill(this.skladZavDBDataSet1.Password);
            DataRow[] UserKey = this.skladZavDBDataSet1.Tables["Password"].Select("User = 'User1'");
            DataRow[] AdminKey = this.skladZavDBDataSet1.Tables["Password"].Select("User = 'Admin'");
            if (textBox1.Text == UserKey[0]["Key"].ToString())
            {
                textBox1.Text = null;
                this.Visible = false;
                Form1 f1 = new Form1();
                f1.Show();
            }
            else
            {
                if (textBox1.Text == AdminKey[0]["Key"].ToString())
                {
                    textBox1.Text = null;
                    this.Visible = false;
                    MessageBox.Show("Выполнен вход с паролем Администратора.", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Form1 f1 = new Form1();
                    f1.Show();
                }
                else
                {
                    if (textBox1.Text == "")
                    {
                        MessageBox.Show("Заполните пустое поле!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        textBox1.Text = null;
                        MessageBox.Show("Неверный пароль!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            try
            {
                this.passwordTableAdapter.Fill(this.skladZavDBDataSet1.Password);
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

        private void Form4_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
            }
        }
    }
}
