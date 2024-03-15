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
using Oracle.ManagedDataAccess.Client;
using System.Configuration;
using DAL;
using Models;

namespace SPM
{
    public partial class GoodIssue : Form
    {
        DateTime time;
        SurplusSolderServic ssServic = new SurplusSolderServic();
        public GoodIssue()
        {
            InitializeComponent();

        }


        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();

            this.Visible = false;
            form1.Visible = true;
            this.Close();


        }

        //private void tabPage1_Click(object sender, EventArgs e)
        //{
        //    MessageBox.Show("tab1 !", "錯誤提示");
        //}
        //private void tabPage2_Click(object sender, EventArgs e)
        //{
        //    MessageBox.Show("tab2 !", "錯誤提示");
        //}

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab.Name == "tabPage2")
                textBox2.Select();
            else if (tabControl1.SelectedTab.Name == "tabPage1")
                textBox3.Select();
        }

        private void GoodIssue_Load(object sender, EventArgs e)
        {
            label6.Text = Class1.PName;//畫面傳值:receive
            label7.Text = Class1.PName1;//畫面傳值:receive
            timer1.Tick += new EventHandler(Timer1_Tick);
            //textBox3.Focus();
            textBox3.Select();
            label8.Text = GetLocQuantity();
            tabControl1.SelectedIndexChanged += tabControl1_SelectedIndexChanged;
            textBox1.KeyDown += textBox1_KeyDown;
            textBox2.KeyDown += textBox2_KeyDown;

        }
        private void Timer1_Tick(object Sender, EventArgs e)
        {
            // Set the caption to the current time.  
            label1.Text = "Date:" + GetDate.Now();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SurplusSolder objSurplusSolder = ssServic.GetSurplusSolder();
            if (objSurplusSolder != null)
            {
                MessageBox.Show("有回庫錫膏未使用，請點擊‘回庫錫膏’按扭");
                return;
            }

            if (textBox3.Text.Trim() == "")
            {
                //Console.Beep(2000, 1000); //2000HZ,1 sec
                MessageBox.Show("請填數量 !", "錯誤提示");
                textBox3.Select();
            }
            else if (Int32.Parse(textBox3.Text) <= 0)
            {
                //Console.Beep(2000, 1000); //2000HZ,1 sec
                MessageBox.Show("請填數量 !", "錯誤提示");
                textBox3.Select();
            }
            else
            {
                //DB connection
                //string constr = "Data Source = 10.2.0.8; User Id = sa; Password = Mes123456; Initial Catalog = GTIMES; Max Pool Size = 300";
                //string constr = "server=10.2.0.230;user id=sa;pwd=ks@12345;database=GTIMES";
                string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
                SqlConnection SqlCon2 = new SqlConnection(ConStr);
                SqlCon2.Open();

                //select DATEDIFF(DAYOFYEAR,TranTime,GETDATE()) from [GTIMES].[dbo].[SolderPaste_Transaction] where Status='1'
                //select top(1) SolderID, Loc from [GTIMES].[dbo].[SolderPaste_Transaction] where Status='1'
                string SqlStr = "select top(" + Int32.Parse(textBox3.Text) + ") a.*,b.* from SolderPaste_Transaction a,SolderPaste_Location b where a.Status='1' " +
                    "and a.SolderID not in(select SolderID from SolderPaste_Transaction where Status = '2' or Status = '3' or Status = '4' or Status = '5' or Status = '6') " +
                    "and a.Loc=b.Loc and b.IsEmpty='N' and " +
                    "DATEDIFF(DAYOFYEAR,a.TranTime,GETDATE())<180  order by trantime,left(a.SolderID,10)";//0302.修改建議儲位排序原則,以入庫日期+DVP前10碼做排序
                DataSet ds = new DataSet();
                SqlDataAdapter Sda = new SqlDataAdapter(SqlStr, SqlCon2);
                Sda.Fill(ds);
                //dataGridView1.DataSource = ds;
                DataTable dt = new DataTable();
                dt.Columns.Add("建議錫膏", typeof(string));
                dt.Columns.Add("取出的儲位", typeof(string));
                dt.Columns.Add("刷入錫膏", typeof(string));
                dt.Columns.Add("開始回溫時間", typeof(string));

                //欄寬自動調整
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                //dataGridView1.EditMode=
                for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                {

                    dt.Rows.Add(ds.Tables[0].Rows[j]["SolderID"].ToString(), ds.Tables[0].Rows[j]["Loc"].ToString());
                }

                dataGridView1.DataSource = dt;
                if (ds.Tables[0].Rows.Count < Int32.Parse(textBox3.Text))
                    MessageBox.Show("可用數量不足 !", "警告提示");
                textBox1.Select();
                SqlCon2.Close();
                SqlCon2.Dispose();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            //
            // Detect the KeyEventArg's key enumerated constant.
            //
            if (e.KeyCode == Keys.Enter)
            {

                string InsStr, UpdStr, InsOra;
                Boolean fg = false;
                string format = "yyyy-MM-dd HH:mm:ss";
                SurplusSolder objSS = ssServic.GetBySolderID(textBox1.Text.Trim());//查询输入的锡膏条码是否为二次回库的
                if (textBox1.TextLength == 28)
                {
                    //DB connection
                    //string constr = "Data Source = 10.2.0.8; User Id = sa; Password = Mes123456; Initial Catalog = GTIMES; Max Pool Size = 300";
                    //string constr = "server=10.2.0.230;user id=sa;pwd=ks@12345;database=GTIMES";
                    string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
                    SqlConnection SqlCon2 = new SqlConnection(ConStr);
                    SqlCon2.Open();


                    //oracle
                    string OracleDB = ConfigurationManager.ConnectionStrings["OracleDB"].ConnectionString; //连接app.config SQL数据库
                    OracleConnection conn = new OracleConnection(OracleDB);
                    //conn.ConnectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.8.200.30)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=GTIMESTS)));Persist Security Info=True;User ID=gtimes;Password=test;";
                    //conn.ConnectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.8.200.30)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=GTIMES)));Persist Security Info=True;User ID=gtimes;Password=systex;";
                    conn.Open();

                    if (objSS != null)
                    {
                        for (int rows = 0; rows < dataGridView1.Rows.Count - 1; rows++)
                        {
                            if (textBox1.Text == dataGridView1.Rows[rows].Cells[0].Value.ToString())
                            {
                                if (textBox1.Text == dataGridView1.Rows[rows].Cells[2].Value.ToString()) //重覆刷入
                                {
                                    fg = false;
                                    break;
                                }

                                InsStr = "insert into SolderPaste_Transaction values ('" + textBox1.Text + "','8','" + GetDate.Now() + "','" + label6.Text + "','','NULL')";


                                //SqlCommand Cmd2 = new SqlCommand(InsStr, SqlCon2);
                                SqlCommand Cmd2 = SqlCon2.CreateCommand();
                                Cmd2.Connection = SqlCon2;
                                Cmd2.CommandText = InsStr;
                                //begin trans
                                SqlTransaction transaction;
                                // Start a local transaction.
                                transaction = SqlCon2.BeginTransaction("SampleTransaction");
                                Cmd2.Transaction = transaction;

                                //Cmd2.ExecuteNonQuery();
                                try
                                {
                                    //寫入SQL錫膏檔
                                    Cmd2.ExecuteNonQuery();

                                    //更改二次回库的备注不F
                                    UpdStr = "update SolderPaste_Transaction set ReMark='F' where ReMark='T' and SolderID='" + textBox1.Text.Trim() + "'";
                                    //Cmd2 = new SqlCommand(UpdStr, SqlCon2);
                                    Cmd2.CommandText = UpdStr;
                                    Cmd2.ExecuteNonQuery();

                                    // Attempt to commit the transaction.
                                    transaction.Commit();

                                    //set value to dataGridView1
                                    dataGridView1.Rows[rows].Cells[2].Value = textBox1.Text;
                                    dataGridView1.Rows[rows].Cells[3].Value = DateTime.Now.ToString(format);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("Commit Exception Type: {0}", ex.GetType());
                                    Console.WriteLine("  Message: {0}", ex.Message);
                                    MessageBox.Show("系統處理失敗 !", "錯誤提示");

                                    // Attempt to roll back the transaction.
                                    try
                                    {
                                        transaction.Rollback();
                                    }
                                    catch (Exception ex2)
                                    {
                                        // This catch block will handle any errors that may have occurred
                                        // on the server that would cause the rollback to fail, such as
                                        // a closed connection.
                                        Console.WriteLine("Rollback Exception Type: {0}", ex2.GetType());
                                        Console.WriteLine("  Message: {0}", ex2.Message);
                                        MessageBox.Show("系統處理失敗 !", "錯誤提示");
                                    }
                                }
                                fg = true;
                                SqlCon2.Close();
                                SqlCon2.Dispose();
                                conn.Close();
                                conn.Dispose();
                            }
                        }
                    }
                    else
                    {
                        for (int rows = 0; rows < dataGridView1.Rows.Count - 1; rows++)
                        {
                            if (textBox1.Text == dataGridView1.Rows[rows].Cells[0].Value.ToString())
                            {
                                if (textBox1.Text == dataGridView1.Rows[rows].Cells[2].Value.ToString()) //重覆刷入
                                {
                                    fg = false;
                                    break;
                                }

                                InsStr = "insert into SolderPaste_Transaction values ('" + textBox1.Text + "','2','"
                                    + GetDate.Now()
                                    + "','" + label6.Text + "','" + dataGridView1.Rows[rows].Cells[1].Value + "','')";

                                //SqlCommand Cmd2 = new SqlCommand(InsStr, SqlCon2);
                                SqlCommand Cmd2 = SqlCon2.CreateCommand();
                                Cmd2.Connection = SqlCon2;
                                Cmd2.CommandText = InsStr;
                                //begin trans
                                SqlTransaction transaction;
                                // Start a local transaction.
                                transaction = SqlCon2.BeginTransaction("SampleTransaction");
                                Cmd2.Transaction = transaction;

                                //Cmd2.ExecuteNonQuery();
                                try
                                {
                                    //寫入SQL錫膏檔
                                    Cmd2.ExecuteNonQuery();

                                    //更新SQL儲位檔
                                    UpdStr = "update SolderPaste_Location set IsEmpty='Y' where Loc=" + dataGridView1.Rows[rows].Cells[1].Value;
                                    //Cmd2 = new SqlCommand(UpdStr, SqlCon2);
                                    Cmd2.CommandText = UpdStr;
                                    Cmd2.ExecuteNonQuery();

                                    // Attempt to commit the transaction.
                                    transaction.Commit();
                                    //Console.WriteLine("Both records are written to database.");

                                    //Oracle MES 錫膏主檔
                                    InsOra = "insert into PF_SOLDER_PASTE values ('" + textBox1.Text + "','" +
                                        textBox1.Text + "',to_date('" + GetDate.Now() + "','yyyy-MM-dd HH24:mi:ss'),'" + label6.Text + "',null,null,null,null)";
                                    OracleCommand oracmd = conn.CreateCommand();
                                    oracmd.CommandText = InsOra;
                                    oracmd.ExecuteNonQuery();

                                    //set value to dataGridView1
                                    dataGridView1.Rows[rows].Cells[2].Value = textBox1.Text;
                                    dataGridView1.Rows[rows].Cells[3].Value = DateTime.Now.ToString(format);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("Commit Exception Type: {0}", ex.GetType());
                                    Console.WriteLine("  Message: {0}", ex.Message);
                                    MessageBox.Show("系統處理失敗 !", "錯誤提示");

                                    // Attempt to roll back the transaction.
                                    try
                                    {
                                        transaction.Rollback();
                                    }
                                    catch (Exception ex2)
                                    {
                                        // This catch block will handle any errors that may have occurred
                                        // on the server that would cause the rollback to fail, such as
                                        // a closed connection.
                                        Console.WriteLine("Rollback Exception Type: {0}", ex2.GetType());
                                        Console.WriteLine("  Message: {0}", ex2.Message);
                                        MessageBox.Show("系統處理失敗 !", "錯誤提示");
                                    }
                                }

                                //end trans

                                ////Oracle MES 錫膏主檔
                                //InsOra = "insert into PF_SOLDER_PASTE values ('" + dataGridView1.Rows[rows].Cells[2].Value + "','" +
                                //    dataGridView1.Rows[rows].Cells[2].Value + "',to_date('" + DateTime.Now.ToString(format) + "','yyyy-MM-dd HH24:mi:ss'),'"+label6.Text+"',null,null,null,null)";
                                //OracleCommand oracmd = conn.CreateCommand();
                                //oracmd.CommandText = InsOra;
                                //oracmd.ExecuteNonQuery();

                                ////更新儲位檔
                                //UpdStr = "update SolderPaste_Location set IsEmpty='Y' where Loc=" + dataGridView1.Rows[rows].Cells[1].Value;

                                //SqlCommand Cmd3 = new SqlCommand(UpdStr, SqlCon2);
                                //Cmd3.ExecuteNonQuery();

                                fg = true;
                                SqlCon2.Close();
                                SqlCon2.Dispose();
                                conn.Close();
                                conn.Dispose();
                            }

                        } //for
                    }
                    if (fg == false)
                    {
                        //Console.Beep(2000, 1000); //2000HZ,1 sec
                        MessageBox.Show("刷錯錫膏 !", "錯誤提示");
                        SqlCon2.Close();
                        SqlCon2.Dispose();
                        conn.Close();
                        conn.Dispose();
                        textBox1.Text = "";
                    }
                } //if
                fg = false;
                textBox1.Text = "";
                label8.Text = GetLocQuantity();
            }

        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {



        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            //
            // Detect the KeyEventArg's key enumerated constant.
            //
            if (e.KeyCode == Keys.Enter)
            {

                int sLoc = 0;
                string InsStr, UpdStr, DelStr;
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
                    //ds.Tables[0].Rows[0]["Status"].ToString() != "2" ||
                    if (objstr != "2" && objstr != "8")//****************************************
                    {
                        //Console.Beep(2000, 1000); //2000HZ,1 sec
                        MessageBox.Show("刷錯錫膏 !", "錯誤提示");
                        SqlCon2.Close();
                        SqlCon2.Dispose();
                        textBox2.Text = "";
                    }
                    if (objstr == "2" || objstr == "8")//**********************
                    {

                        DataTable dt = new DataTable();
                        dt.Columns.Add("錫膏 DVP", typeof(string));
                        dt.Columns.Add("開始回溫時間", typeof(string));
                        dt.Columns.Add("回溫取消時間", typeof(string));
                        dt.Columns.Add("入冰箱儲位", typeof(string));

                        //欄寬自動調整
                        dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                        for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                        {

                            dt.Rows.Add(ds.Tables[0].Rows[j]["SolderID"].ToString(), ds.Tables[0].Rows[j]["TranTime"].ToString());
                        }

                        dataGridView2.DataSource = dt;
                        if (objstr == "8")//****************************
                        {
                            for (int rows = 0; rows < dataGridView2.Rows.Count - 1; rows++)
                            {
                                if (textBox2.Text == dataGridView2.Rows[rows].Cells[0].Value.ToString())
                                {
                                    //set value to dataGridView
                                    dataGridView2.Rows[rows].Cells[2].Value = DateTime.Now.ToString(format);

                                    //刪除交易檔
                                    DelStr = "delete from SolderPaste_Transaction where SolderID='" + textBox2.Text + "' and Status='8'";
                                    SqlCommand Cmd1 = new SqlCommand(DelStr, SqlCon2);
                                    Cmd1.ExecuteNonQuery();


                                    //string SqlStr1 = "select top(1) Loc from SolderPaste_Transaction where SolderID='" + textBox2.Text + "' order by TranTime  desc";
                                    //DataSet ds1 = new DataSet();
                                    //SqlDataAdapter Sda1 = new SqlDataAdapter(SqlStr1, SqlCon2);
                                    //Sda1.Fill(ds1);
                                    //for (int j = 0; j < ds1.Tables[0].Rows.Count; j++)
                                    //{
                                    //    sLoc = Int32.Parse(ds1.Tables[0].Rows[j]["Loc"].ToString());
                                    //    dataGridView2.Rows[rows].Cells[3].Value = sLoc.ToString();
                                    //}
                                    //end of getting Loc

                                    //UpdStr = "update SolderPaste_Transaction set TranTime='" + dataGridView2.Rows[rows].Cells[2].Value + "',UserID='" + label6.Text + "',Loc="+ sLoc +" where SolderID='"+ textBox2.Text+"'";
                                    UpdStr = "update SolderPaste_Transaction set TranTime='" + dataGridView2.Rows[rows].Cells[2].Value + "',UserID='" + label6.Text + "' where SolderID='" + textBox2.Text + "'";
                                    SqlCommand Cmd2 = new SqlCommand(UpdStr, SqlCon2);
                                    Cmd2.ExecuteNonQuery();

                                    //更新备注
                                    UpdStr = "update SolderPaste_Transaction set ReMark='T' where ReMark='F' and status='7' and SolderID='" + textBox2.Text + "'";
                                    SqlCommand Cmd21 = new SqlCommand(UpdStr, SqlCon2);
                                    Cmd21.ExecuteNonQuery();

                                    //新增Log檔
                                    InsStr = "insert into SolderPaste_CancelLog values('" + textBox2.Text + "','8','" + dataGridView2.Rows[rows].Cells[2].Value + "')";
                                    SqlCommand Cmd3 = new SqlCommand(InsStr, SqlCon2);
                                    Cmd3.ExecuteNonQuery();

                                    //以下Lief新增 for 入冰箱
                                    //UpdStr = "update SolderPaste_Transaction set TranTime='" + dataGridView2.Rows[rows].Cells[2].Value + "',UserID='" + label6.Text + "' where SolderID='" + textBox2.Text + "' and Status='1' ";
                                    //SqlCommand Cmd_lief = new SqlCommand(UpdStr, SqlCon2);
                                    //Cmd_lief.ExecuteNonQuery();
                                }

                            } //for
                        }

                        //===========
                        if (objstr == "2")//****************************
                        {
                            for (int rows = 0; rows < dataGridView2.Rows.Count - 1; rows++)
                            {
                                if (textBox2.Text == dataGridView2.Rows[rows].Cells[0].Value.ToString())
                                {
                                    //set value to dataGridView
                                    dataGridView2.Rows[rows].Cells[2].Value = DateTime.Now.ToString(format);

                                    //刪除交易檔
                                    DelStr = "delete from SolderPaste_Transaction where SolderID='" + textBox2.Text + "' and Status='2'";
                                    SqlCommand Cmd1 = new SqlCommand(DelStr, SqlCon2);
                                    Cmd1.ExecuteNonQuery();

                                    //更新交易檔
                                    //get Loc                        
                                    //string SqlStr1 = "select top(1) Loc from SolderPaste_Location where IsEmpty='Y' order by Loc ASC";
                                    //DataSet ds1 = new DataSet();
                                    //SqlDataAdapter Sda1 = new SqlDataAdapter(SqlStr1, SqlCon2);
                                    //Sda1.Fill(ds1);
                                    //for (int j = 0; j < ds1.Tables[0].Rows.Count; j++)
                                    //{
                                    //    sLoc = Int32.Parse(ds1.Tables[0].Rows[j]["Loc"].ToString());
                                    //    dataGridView2.Rows[rows].Cells[3].Value = sLoc.ToString();
                                    //}
                                    string SqlStr1 = "select top(1) Loc from SolderPaste_Transaction where SolderID='" + textBox2.Text + "' order by TranTime  desc";
                                    DataSet ds1 = new DataSet();
                                    SqlDataAdapter Sda1 = new SqlDataAdapter(SqlStr1, SqlCon2);
                                    Sda1.Fill(ds1);
                                    for (int j = 0; j < ds1.Tables[0].Rows.Count; j++)
                                    {
                                        sLoc = Int32.Parse(ds1.Tables[0].Rows[j]["Loc"].ToString());
                                        dataGridView2.Rows[rows].Cells[3].Value = sLoc.ToString();
                                    }
                                    //end of getting Loc

                                    //UpdStr = "update SolderPaste_Transaction set TranTime='" + dataGridView2.Rows[rows].Cells[2].Value + "',UserID='" + label6.Text + "',Loc="+ sLoc +" where SolderID='"+ textBox2.Text+"'";
                                    UpdStr = "update SolderPaste_Transaction set TranTime='" + dataGridView2.Rows[rows].Cells[2].Value + "',UserID='" + label6.Text + "' where SolderID='" + textBox2.Text + "'";
                                    SqlCommand Cmd2 = new SqlCommand(UpdStr, SqlCon2);
                                    Cmd2.ExecuteNonQuery();

                                    //更新儲位檔
                                    UpdStr = "update SolderPaste_Location set IsEmpty='N' where Loc=" + sLoc;
                                    SqlCommand Cmd21 = new SqlCommand(UpdStr, SqlCon2);
                                    Cmd21.ExecuteNonQuery();

                                    //新增Log檔
                                    InsStr = "insert into SolderPaste_CancelLog values('" + textBox2.Text + "','2','" + dataGridView2.Rows[rows].Cells[2].Value + "')";
                                    SqlCommand Cmd3 = new SqlCommand(InsStr, SqlCon2);
                                    Cmd3.ExecuteNonQuery();

                                    //以下Lief新增 for 入冰箱
                                    //UpdStr = "update SolderPaste_Transaction set TranTime='" + dataGridView2.Rows[rows].Cells[2].Value + "',UserID='" + label6.Text + "' where SolderID='" + textBox2.Text + "' and Status='1' ";
                                    //SqlCommand Cmd_lief = new SqlCommand(UpdStr, SqlCon2);
                                    //Cmd_lief.ExecuteNonQuery();
                                }

                            } //for
                        }

                        //===========
                        textBox2.Text = "";
                        textBox2.Select();

                        SqlCon2.Close();
                        SqlCon2.Dispose();
                    }

                } //if length
            } //if keycode

        }


        private void textBox2_TextChanged(object sender, EventArgs e)
        {
        }

        public static string GetLocQuantity() //取得冰箱可用的鍚膏數量
        {
            string str = "";

            //連接sql
            string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
            SqlConnection SqlCon = new SqlConnection(ConStr);
            SqlCon.Open();

            //宣告DataReader
            SqlDataReader dr = null;
            SqlCommand Cmd = new SqlCommand("Select isnull(count(*),0) as count from SolderPaste_Location with(nolock) where isempty = 'N'", SqlCon);
            dr = Cmd.ExecuteReader(); //執行並回傳

            if (dr.HasRows)
            {
                while (dr.Read())
                {
                    str = dr["count"].ToString();

                }
            }

            Cmd.Cancel();
            dr.Close();
            SqlCon.Close();
            SqlCon.Dispose();
            return str;
        }

        private void btnSurplusSolder_Click(object sender, EventArgs e)
        {
            string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
            SqlConnection SqlCon2 = new SqlConnection(ConStr);
            SqlCon2.Open();

            string SqlStr = "select SolderID from SolderPaste_Transaction where Status = '7' and ReMark = 'T'";
            DataSet ds = new DataSet();
            SqlDataAdapter Sda = new SqlDataAdapter(SqlStr, SqlCon2);
            Sda.Fill(ds);
            //dataGridView1.DataSource = ds;
            DataTable dt = new DataTable();
            dt.Columns.Add("建議錫膏", typeof(string));
            dt.Columns.Add("取出的儲位", typeof(string));
            dt.Columns.Add("刷入錫膏", typeof(string));
            dt.Columns.Add("開始回溫時間", typeof(string));

            //欄寬自動調整
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            //dataGridView1.EditMode=
            for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
            {

                //dt.Rows.Add(ds.Tables[0].Rows[j]["SolderID"].ToString(), ds.Tables[0].Rows[j]["Loc"].ToString());
                dt.Rows.Add(ds.Tables[0].Rows[j]["SolderID"].ToString());
            }

            dataGridView1.DataSource = dt;
            //if (ds.Tables[0].Rows.Count < Int32.Parse(textBox3.Text))
            //    MessageBox.Show("可用數量不足 !", "警告提示");
            textBox1.Select();
            SqlCon2.Close();
            SqlCon2.Dispose();
        }
    }
}
