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
    public partial class QueryPerSolder : Form
    {
        public QueryPerSolder()
        {
            InitializeComponent();
        }

        private void QueryPerSolder_Load(object sender, EventArgs e)
        {
            label6.Text = Class1.PName;//畫面傳值:receive
            label7.Text = Class1.PName1;//畫面傳值:receive
            textBox2.KeyDown += textBox2_KeyDown;
            textBox2.Select();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();

            this.Visible = false;
            form1.Visible = true;
            this.Close();
        }

        

        private void textBox2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            //
            // Detect the KeyEventArg's key enumerated constant.
            //
            if (e.KeyCode == Keys.Enter)
            {

                if (textBox2.TextLength == 28)
                {
                    label1.Text = "";
                    //DB connection
                    //string constr = "Data Source = 10.2.0.8; User Id = sa; Password = Mes123456; Initial Catalog = GTIMES; Max Pool Size = 300";
                    //string constr = "server=10.2.0.230;user id=sa;pwd=ks@12345;database=GTIMES";
                    string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
                    SqlConnection SqlCon2 = new SqlConnection(ConStr);
                    SqlCon2.Open();

                    string SqlStr = "select * from SolderPaste_Transaction where SolderID='" + textBox2.Text + "' order by TranTime DESC";
                    DataSet ds = new DataSet();
                    SqlDataAdapter Sda = new SqlDataAdapter(SqlStr, SqlCon2);
                    Sda.Fill(ds);

                    if (ds.Tables[0].Rows.Count <= 0)
                    {
                        //Console.Beep(2000, 1000); //2000HZ,1 sec
                        MessageBox.Show("錫膏無狀態 !", "錯誤提示");
                        SqlCon2.Close();
                        SqlCon2.Dispose();
                        textBox2.Text = "";
                        //label1.Text = "";
                    }
                    else
                    {
                        switch (ds.Tables[0].Rows[0]["Status"].ToString())
                        {
                            case "1":
                                label1.Text = "入庫";
                                break;
                            case "2":
                                label1.Text = "回溫";
                                break;
                            case "3":
                                label1.Text = "攪拌";
                                break;
                            case "4":
                                label1.Text = "使用";
                                break;
                            case "5":
                                label1.Text = "結束";
                                break;
                            case "6":
                                label1.Text = "報廢";
                                break;
                            case "7":
                                label1.Text = "二次回庫";
                                break;
                            case "8":
                                label1.Text = "二次回溫";
                                break;
                            case "9":
                                label1.Text = "二次攪拌";
                                break;
                            case "A":
                                label1.Text = "二次使用";
                                break;
                            default:

                                break;
                        }



                        //===========
                        textBox2.Text = "";
                        textBox2.Select();

                        SqlCon2.Close();
                        SqlCon2.Dispose();
                    }
                } //if length
            }

        }

        

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
        
        }
    }
}
