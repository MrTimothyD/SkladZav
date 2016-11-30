using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SkladZav
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        bool t;
        int temp;
        decimal temp1;

        public static bool Input(out int a, out decimal b)
        {
            Form2 IBform = new Form2();
            IBform.ShowDialog();
            a = IBform.temp;
            b = IBform.temp1;
            return IBform.t;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                temp = Convert.ToInt32(this.textBox1.Text);
                temp1 = Convert.ToDecimal(this.textBox2.Text);
                textBox1.Text = null;
                textBox2.Text = null;
                this.Visible = false;
                t = true;
            }
            catch (Exception)
            {
                textBox1.Text = null;
                textBox2.Text = null;
                MessageBox.Show("Введены неверные значения!", "Заявка по складу", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = null;
            textBox2.Text = null;
            this.Visible = false;
            t = false;            
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8 && number != 44) // цифры, клавиша BackSpace и запятая
            {
                e.Handled = true;
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
            }
        }
    }
}
