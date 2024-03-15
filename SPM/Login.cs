using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace SPM
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void Login_Load(object sender, EventArgs e)
        {
            textBox1.Select();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            //DB connection
            //string constr = "Data Source = 10.2.0.8; User Id = sa; Password = Mes123456; Initial Catalog = GTIMES; Max Pool Size = 300";
            //string constr = "server=10.2.0.230;user id=sa;pwd=ks@12345;database=GTIMES";
            string constr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
            SqlConnection SqlCon2 = new SqlConnection(constr);
            SqlCon2.Open();

            //select DATEDIFF(DAYOFYEAR,TranTime,GETDATE()) from [GTIMES].[dbo].[SolderPaste_Transaction] where Status='1'
            //select top(1) SolderID, Loc from [GTIMES].[dbo].[SolderPaste_Transaction] where Status='1'
            string SqlStr = "select * from SolderPaste_UserMaster where UserID='" + textBox1.Text + "' and Pwd='" + textBox2.Text + "' collate Chinese_Taiwan_Stroke_CS_AS";
            DataSet ds = new DataSet();
            SqlDataAdapter Sda = new SqlDataAdapter(SqlStr, SqlCon2);
            Sda.Fill(ds);
            
            if (ds.Tables[0].Rows.Count<=0)
            {
                //Console.Beep(2000, 1000); //2000HZ,1 sec
                MessageBox.Show("工號或密碼錯誤 !", "錯誤提示");                
                textBox1.Select();
                SqlCon2.Close();
                SqlCon2.Dispose();
            }
            else
            {
                Form1 Form1 = new Form1();
                Class1.PName = textBox1.Text;//畫面傳值:assign
                Class1.PName1 = ds.Tables[0].Rows[0]["Name"].ToString();//畫面傳值:assign
                this.Visible = false;
                Form1.Show();
            }
            
        }
    }
}
