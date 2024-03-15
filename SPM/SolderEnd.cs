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
    public partial class SolderEnd : Form
    {
        public SolderEnd()
        {
            InitializeComponent();
        }

        private void SolderEnd_Load(object sender, EventArgs e)
        {
            label6.Text = Class1.PName;//畫面傳值:receive
            label7.Text = Class1.PName1;//畫面傳值:receive
            textBox2.Select();
            textBox1.KeyDown += textBox1_KeyDown;
            textBox2.KeyDown += textBox2_KeyDown;
            tabControl1.SelectedIndexChanged += tabControl1_SelectedIndexChanged;
        }

        private void textBox1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            //
            // Detect the KeyEventArg's key enumerated constant.
            //
            if (e.KeyCode == Keys.Enter)
            {

                int sLoc = 0;
                string InsStr, UpdStr, DelStr;
                string format = "yyyy-MM-dd HH:mm:ss";
                if (textBox1.TextLength == 28)
                {
                    //DB connection
                    //string constr = "Data Source = 10.2.0.8; User Id = sa; Password = Mes123456; Initial Catalog = GTIMES; Max Pool Size = 300";
                    //string constr = "server=10.2.0.230;user id=sa;pwd=ks@12345;database=GTIMES";
                    string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
                    SqlConnection SqlCon2 = new SqlConnection(ConStr);
                    SqlCon2.Open();

                    string SqlStr = "select top 1 * from SolderPaste_Transaction where SolderID='" + textBox1.Text + "' order by TranTime DESC";
                    DataSet ds = new DataSet();
                    SqlDataAdapter Sda = new SqlDataAdapter(SqlStr, SqlCon2);
                    Sda.Fill(ds);
                    DataTable dt = new DataTable();
                    dt.Columns.Add("錫膏 DVP", typeof(string));
                    dt.Columns.Add("結束時間", typeof(string));
                    dt.Columns.Add("取消結束時間", typeof(string));

                    //欄寬自動調整
                    dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    if (ds.Tables[0].Rows.Count <= 0)
                    {
                        //Console.Beep(2000, 1000); //2000HZ,1 sec
                        MessageBox.Show("刷錯錫膏 !", "錯誤提示");
                        SqlCon2.Close();
                        SqlCon2.Dispose();
                        textBox1.Text = "";
                        textBox1.Select();
                    }
                    if (ds.Tables[0].Rows[0]["Status"].ToString() != "5" )
                    {
                        //Console.Beep(2000, 1000); //2000HZ,1 sec
                        MessageBox.Show("刷錯錫膏 !", "錯誤提示");
                        SqlCon2.Close();
                        SqlCon2.Dispose();
                        textBox1.Text = "";
                        textBox1.Select();
                    }
                    if (ds.Tables[0].Rows[0]["Status"].ToString() == "5")
                    {
                        for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                        {

                            dt.Rows.Add(ds.Tables[0].Rows[j]["SolderID"].ToString(), ds.Tables[0].Rows[j]["TranTime"].ToString(), GetDate.Now());
                        }
                        dataGridView2.DataSource = dt;
                        //===========
                        for (int rows = 0; rows < dataGridView2.Rows.Count - 1; rows++)
                        {
                            if (textBox1.Text == dataGridView2.Rows[rows].Cells[0].Value.ToString())
                            {
                                //刪除交易檔
                                DelStr = "delete from SolderPaste_Transaction where SolderID='" + textBox1.Text + "' and Status='5'";
                                SqlCommand Cmd1 = new SqlCommand(DelStr, SqlCon2);
                                Cmd1.ExecuteNonQuery();

                                //新增Log檔
                                InsStr = "insert into SolderPaste_CancelLog values('" + textBox1.Text + "','5','" + GetDate.Now() + "')";
                                SqlCommand Cmd3 = new SqlCommand(InsStr, SqlCon2);
                                Cmd3.ExecuteNonQuery();
                            }

                        } //for read gridview

                        //===========
                        textBox1.Text = "";
                        textBox1.Select();

                        SqlCon2.Close();
                        SqlCon2.Dispose();
                    }
                } //if length
            }

        }


        private void textBox2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            //
            // Detect the KeyEventArg's key enumerated constant.
            //
            if (e.KeyCode == Keys.Enter)
            {

                string InsStr, UpdStr;
                Boolean fg = false;
                string format = "yyyy-MM-dd HH:mm:ss";
                if (textBox2.TextLength == 28)
                {
                    //DB connection
                    //string constr = "Data Source = 10.2.0.8; User Id = sa; Password = Mes123456; Initial Catalog = GTIMES; Max Pool Size = 300";
                    //string constr = "server=10.2.0.230;user id=sa;pwd=ks@12345;database=GTIMES";
                    string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
                    SqlConnection SqlCon2 = new SqlConnection(ConStr);
                    SqlCon2.Open();
                    string SqlStr = "select top 1 * from SolderPaste_Transaction where SolderID='" + textBox2.Text + "' order by TranTime DESC";
                    DataSet ds = new DataSet();
                    SqlDataAdapter Sda = new SqlDataAdapter(SqlStr, SqlCon2);
                    Sda.Fill(ds);

                    string objstr = ds.Tables[0].Rows[0]["Status"].ToString();
                    //dataGridView1.DataSource = ds;
                    DataTable dt = new DataTable();
                    dt.Columns.Add("錫膏 DVP", typeof(string));
                    dt.Columns.Add("使用時間", typeof(string));
                    dt.Columns.Add("結束時間", typeof(string));

                    //欄寬自動調整
                    dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    if (ds.Tables[0].Rows.Count <= 0)
                    {
                        //Console.Beep(2000, 1000); //2000HZ,1 sec
                        MessageBox.Show("刷錯錫膏 !", "錯誤提示");
                        SqlCon2.Close();
                        SqlCon2.Dispose();
                        textBox2.Text = "";
                    }
                    if (objstr != "4" && objstr != "A")//4为正常使用，10为二次使用
                    {
                        //Console.Beep(2000, 1000); //2000HZ,1 sec
                        MessageBox.Show("刷錯錫膏 !", "錯誤提示");
                        SqlCon2.Close();
                        SqlCon2.Dispose();
                        textBox2.Text = "";
                    }
                    if (objstr == "4" || objstr == "A")//ds.Tables[0].Rows[0]["Status"].ToString()
                    {
                        for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                        {

                            dt.Rows.Add(ds.Tables[0].Rows[j]["SolderID"].ToString(), ds.Tables[0].Rows[j]["TranTime"].ToString(), GetDate.Now());
                        }
                        dataGridView1.DataSource = dt;

                        textBox2.Select();
                        //SqlCon2.Close();
                        //SqlCon2.Dispose();

                        for (int rows = 0; rows < dataGridView1.Rows.Count - 1; rows++)
                        {
                            if (textBox2.Text == dataGridView1.Rows[rows].Cells[0].Value.ToString())
                            {
                                //set value to dataGridView1
                                //dataGridView1.Rows[rows].Cells[2].Value = textBox2.Text;
                                //dataGridView1.Rows[rows].Cells[3].Value = DateTime.Now.ToString(format);

                                InsStr = "insert into SolderPaste_Transaction values ('" + textBox2.Text + "','5','"
                                    + dataGridView1.Rows[rows].Cells[2].Value
                                    + "','" + label6.Text + "','" + ds.Tables[0].Rows[0]["Loc"].ToString() + "','')";

                                SqlCommand Cmd2 = new SqlCommand(InsStr, SqlCon2);
                                Cmd2.ExecuteNonQuery();

                                textBox2.Text = "";
                                SqlCon2.Close();
                                SqlCon2.Dispose();
                            }

                        } //for

                    } //if
                } //if
            }

        }


        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab.Name == "tabPage2")
                textBox1.Select();
            else if (tabControl1.SelectedTab.Name == "tabPage1")
                textBox2.Select();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();

            this.Visible = false;
            form1.Visible = true;
            this.Close();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        } //void

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

    }
}
