using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SPM
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Start();
            label6.Text = Class1.PName;
            label5.Text = Class1.PName1;
        }

        private void date_Now(object sender, EventArgs e)
        {
            label1.Text =GetDate.Now();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();

            this.Visible = false;
            form2.Visible = true;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form4 form4 = new Form4();

            this.Visible = false;
            form4.Visible = true;

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            GoodIssue GoodIssue = new GoodIssue();
            //Class1.PName = label6.Text;//畫面傳值:assign
            this.Visible = false;
            //GoodIssue.Visible = true;
            GoodIssue.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SolderEnd SolderEnd = new SolderEnd();
            //Class1.PName = label6.Text;//畫面傳值:assign
            this.Visible = false;
            //GoodIssue.Visible = true;
            SolderEnd.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form5 form5 = new Form5();

            this.Visible = false;
            form5.Visible = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            QueryPerSolder QueryPerSolder = new QueryPerSolder();
            this.Visible = false;
            QueryPerSolder.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            QueryHistory QueryHistory = new QueryHistory();
            this.Visible = false;
            QueryHistory.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            
            QueryStock QueryStock = new QueryStock();
            this.Visible = false;
            QueryStock.Show();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            FrmSurplusSolder surplusSolder = new FrmSurplusSolder();
            this.Visible = false;
            surplusSolder.Show();
            
        }
    }
}
