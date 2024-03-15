using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Oracle.ManagedDataAccess.Client;
using System.Configuration;

namespace SPM
{
    public partial class Form5 : Form
    {
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dt_show1 = new DataTable();

        public Form5()
        {
            InitializeComponent();
            this.ControlBox = false;
        }

        private void Form5_Load(object sender, EventArgs e)
        {
            timer1.Start();
            label6.Text = Class1.PName;//畫面傳值:receive
            label7.Text = Class1.PName1;//畫面傳值:receive
            textBox1.KeyDown += textBox1_KeyDown;
            textBox2.KeyDown += textBox2_KeyDown;
            dt.Columns.Add("DVP", typeof(string));
            dt.Columns.Add("Status", typeof(string));
            dt.Columns.Add("TranTime", typeof(DateTime));
            dt.Columns.Add("UserID", typeof(string));
            dt.Columns.Add("Loc", typeof(string));

            dt_show1.Columns.Add("DVP", typeof(string));
            dt_show1.Columns.Add("已攪拌時間", typeof(string));
            dt_show1.Columns.Add("開始使用時間", typeof(DateTime));

            dt2.Columns.Add("DVP", typeof(string));

            //放入dataGridView顯示、欄位寬度自動調整      
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.DataSource = dt_show1;

            textBox1.Select();
            tabControl1.SelectedIndexChanged += tabControl1_SelectedIndexChanged;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab.Name == "tabPage2")
            {
                textBox2.Select();
                dt.Clear();
                dt_show1.Clear();
                dataGridView1.DataSource = dt_show1;
            }
            if (tabControl1.SelectedTab.Name == "tabPage1")
            {
                textBox1.Select();
                dt2.Clear();
                dataGridView2.DataSource = dt2;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            this.Visible = false;
            form1.Visible = true;
            this.Close();

        }

        private void textBox1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            //
            // Detect the KeyEventArg's key enumerated constant.
            //
            if (e.KeyCode == Keys.Enter)
            {

                if (textBox1.TextLength == 28)
                {
                    string id = this.textBox1.Text;
                    string check = "";
                    string loc = "";
                    DateTime dateTime_now = Convert.ToDateTime(GetDate.Now());//2021.04.21修改程式使用服務器時間來卡控
                    DateTime dateTime = Convert.ToDateTime(GetDate.Now());

                    if (id.Length == 28)
                    {
                        try
                        {
                            #region 確認該錫膏的最新狀態

                            //連接sql
                            //string ConStr = "server=10.2.0.8;user id=sa;pwd=Mes123456;database=GTIMES";
                            //string ConStr = "server=10.2.0.230;user id=sa;pwd=ks@12345;database=GTIMES";
                            string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
                            SqlConnection SqlCon = new SqlConnection(ConStr);

                            SqlCon.Open();  //open

                            //宣告DataReader
                            SqlDataReader dr = null;
                            SqlCommand Cmd_item = new SqlCommand("select  top 1 status as status,TranTime,loc from SolderPaste_Transaction where SolderID='" + id + "' order by TranTime  DESC", SqlCon);
                            dr = Cmd_item.ExecuteReader(); //執行並回傳

                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    check = dr["status"].ToString();
                                    string str_time = dr["TranTime"].ToString();
                                    dateTime = Convert.ToDateTime(str_time);
                                    loc = dr["loc"].ToString();

                                }

                            }

                            Cmd_item.Cancel(); //關閉DataReader前先關閉sqlcommend
                            dr.Close(); //關閉DataReader

                            SqlCon.Close(); //close
                            #endregion

                            TimeSpan timeSpan = dateTime_now - dateTime;
                            int x = Convert.ToInt32(timeSpan.TotalMinutes);

                            if (check == "3" && x > 5) //確認狀態是攪拌且攪拌超過5分鐘
                            {
                                Boolean fg = false; //Weaver
                                dt_show1.Rows.Add(id, timeSpan.Hours.ToString() + "小時" + timeSpan.Minutes.ToString() + "分鐘" + timeSpan.Seconds.ToString() + "秒", dateTime_now);
                                dataGridView1.DataSource = dt_show1;

                                //將紀錄寫入dt
                                dt.Rows.Add(id, "4", GetDate.Now(), label6.Text, loc);

                                #region datatable資料寫入DB

                                SqlCon.Open();
                                using (SqlTransaction sqlTransaction = SqlCon.BeginTransaction())
                                {
                                    using (SqlBulkCopy SBC = new SqlBulkCopy(SqlCon, SqlBulkCopyOptions.Default, sqlTransaction))
                                    //using (SqlBulkCopy SBC = new SqlBulkCopy(SqlCon))
                                    {
                                        SBC.DestinationTableName = "SolderPaste_Transaction"; //table name

                                        //mapping
                                        SBC.ColumnMappings.Add("DVP", "SolderID");
                                        SBC.ColumnMappings.Add("Status", "Status");
                                        SBC.ColumnMappings.Add("TranTime", "TranTime");
                                        SBC.ColumnMappings.Add("UserID", "UserID");
                                        SBC.ColumnMappings.Add("Loc", "Loc");

                                        try
                                        {
                                            SBC.WriteToServer(dt); //寫入DB
                                            sqlTransaction.Commit();
                                            fg = true;
                                            SBC.Close();

                                        }
                                        catch
                                        {
                                            sqlTransaction.Rollback();
                                            fg = false;
                                            SBC.Close();
                                            MessageBox.Show("系統處理失敗 !", "錯誤提示");
                                        }
                                    }

                                } //BeginTrans
                                if (fg == true)
                                {
                                    //update oracle MES 錫膏主檔
                                    string UpdOra;
                                    string format = "yyyy-MM-dd HH:mm:ss";
                                    string OracleDB = ConfigurationManager.ConnectionStrings["OracleDB"].ConnectionString; //连接app.config SQL数据库
                                    OracleConnection conn = new OracleConnection(OracleDB);
                                    //conn.ConnectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.8.200.30)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=GTIMESTS)));Persist Security Info=True;User ID=gtimes;Password=test;";
                                    //conn.ConnectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.8.200.30)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=GTIMES)));Persist Security Info=True;User ID=gtimes;Password=systex;";
                                    conn.Open();

                                    OracleCommand oracmd = conn.CreateCommand();
                                    UpdOra = "update PF_SOLDER_PASTE set open_date = to_date('" + GetDate.Now() + "','yyyy-MM-dd HH24:mi:ss'),open_user='" + label6.Text + "' where sp_lot = '" + id + "'";
                                    oracmd.CommandText = UpdOra;
                                    oracmd.ExecuteNonQuery();
                                    conn.Close();
                                    conn.Dispose();
                                }

                                #endregion
                                dt.Clear();
                                this.textBox1.Text = "";

                            }

                            else if (check == "9" && x > 5) //確認狀態是攪拌且攪拌超過5分鐘
                            {
                                Boolean fg = false; //Weaver
                                dt_show1.Rows.Add(id, timeSpan.Hours.ToString() + "小時" + timeSpan.Minutes.ToString() + "分鐘" + timeSpan.Seconds.ToString() + "秒", dateTime_now);
                                dataGridView1.DataSource = dt_show1;

                                //將紀錄寫入dt
                                dt.Rows.Add(id,"A", GetDate.Now(), label6.Text);//A为使用

                                #region datatable資料寫入DB

                                SqlCon.Open();
                                using (SqlTransaction sqlTransaction = SqlCon.BeginTransaction())
                                {
                                    using (SqlBulkCopy SBC = new SqlBulkCopy(SqlCon, SqlBulkCopyOptions.Default, sqlTransaction))
                                    //using (SqlBulkCopy SBC = new SqlBulkCopy(SqlCon))
                                    {
                                        SBC.DestinationTableName = "SolderPaste_Transaction"; //table name

                                        //mapping
                                        SBC.ColumnMappings.Add("DVP", "SolderID");
                                        SBC.ColumnMappings.Add("Status", "Status");
                                        SBC.ColumnMappings.Add("TranTime", "TranTime");
                                        SBC.ColumnMappings.Add("UserID", "UserID");
                                       // SBC.ColumnMappings.Add("Loc", "Loc");

                                        try
                                        {
                                            SBC.WriteToServer(dt); //寫入DB
                                            sqlTransaction.Commit();
                                            fg = true;
                                            SBC.Close();

                                        }
                                        
                                        catch (Exception ex)
                                        {
                                            sqlTransaction.Rollback();
                                            fg = false;
                                            SBC.Close();
                                            MessageBox.Show(ex.Message);
                                            MessageBox.Show("系統處理失敗 !", "錯誤提示");
                                        }
                                    }

                                } //BeginTrans   不写入MES锡膏表中
                                //if (fg == true)
                                //{
                                //    //update oracle MES 錫膏主檔
                                //    string UpdOra;
                                //    string format = "yyyy-MM-dd HH:mm:ss";
                                //    string OracleDB = ConfigurationManager.ConnectionStrings["OracleDB"].ConnectionString; //连接app.config SQL数据库
                                //    OracleConnection conn = new OracleConnection(OracleDB);
                                //    //conn.ConnectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.8.200.30)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=GTIMESTS)));Persist Security Info=True;User ID=gtimes;Password=test;";
                                //    //conn.ConnectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.8.200.30)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=GTIMES)));Persist Security Info=True;User ID=gtimes;Password=systex;";
                                //    conn.Open();

                                //    OracleCommand oracmd = conn.CreateCommand();
                                //    UpdOra = "update PF_SOLDER_PASTE set open_date = to_date('" + GetDate.Now() + "','yyyy-MM-dd HH24:mi:ss'),open_user='" + label6.Text + "' where sp_lot = '" + id + "'";
                                //    oracmd.CommandText = UpdOra;
                                //    oracmd.ExecuteNonQuery();
                                //    conn.Close();
                                //    conn.Dispose();
                                //}

                                #endregion
                                dt.Clear();
                                this.textBox1.Text = "";

                            }
                            else
                            {
                                if (check != "3" && check != "9")//3为正常搅拌，9为二次搅拌
                                {
                                    //Console.Beep(2000, 1000);
                                    MessageBox.Show("此錫膏現在狀態為:" + codemapping(check));
                                    this.textBox1.Text = "";
                                }
                                if (x < 5 && check == "3"|| x < 5 && check == "9")
                                {
                                    //Console.Beep(2000, 1000);
                                    MessageBox.Show("攪拌時間不足");
                                    this.textBox1.Text = "";
                                }
                            }

                        }
                        catch (Exception msg)
                        {
                            //Console.Beep(2000, 1000);
                            MessageBox.Show(msg.ToString());
                            this.textBox1.Text = "";
                        }
                    }
                }
            }

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
                    string id = this.textBox2.Text;
                    string check = "";
                    string loc = "";
                    DateTime dateTime_now = DateTime.Now;
                    DateTime dateTime = DateTime.Now;
                    string time = "";

                    if (id.Length == 28)
                    {
                        try
                        {
                            #region 確認該錫膏的最新狀態

                            //連接sql
                            //string ConStr = "server=10.2.0.8;user id=sa;pwd=Mes123456;database=GTIMES";
                            //string ConStr = "server=10.2.0.230;user id=sa;pwd=ks@12345;database=GTIMES";
                            string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
                            SqlConnection SqlCon = new SqlConnection(ConStr);

                            SqlCon.Open();  //open

                            //宣告DataReader
                            SqlDataReader dr = null;
                            SqlCommand Cmd_item = new SqlCommand("select  top 1 status as status,TranTime,loc from SolderPaste_Transaction where SolderID='" + id + "' order by TranTime  DESC", SqlCon);
                            dr = Cmd_item.ExecuteReader(); //執行並回傳

                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    check = dr["status"].ToString();
                                    dateTime = Convert.ToDateTime(dr["TranTime"].ToString());
                                    time = dateTime.ToString("yyyy-MM-dd hh:mm:ss.fff");
                                    loc = dr["loc"].ToString();

                                }

                            }

                            Cmd_item.Cancel(); //關閉DataReader前先關閉sqlcommend
                            dr.Close(); //關閉DataReader
                            SqlCon.Close(); //close
                            #endregion

                            if (check == "4"|| check == "A") //確認狀態是使用狀態,A为二次使用状态
                            {

                                //刪除最新一筆使用紀錄、寫入log

                                SqlCon.Open();

                                string InsStr = "insert into SolderPaste_CancelLog values ('" + id + "','" + check + "',CONVERT(DATETIME,'" +GetDate.Now()+ "'))";
                                SqlCommand Cmd = new SqlCommand(InsStr, SqlCon);
                                Cmd.ExecuteNonQuery();

                                InsStr = "delete SolderPaste_Transaction  where SolderID='" + id + "' and TranTime=(select  top 1 TranTime from SolderPaste_Transaction where SolderID='" + id + "' order by TranTime  DESC)";
                                Cmd = new SqlCommand(InsStr, SqlCon);
                                Cmd.ExecuteNonQuery();

                                SqlCon.Close();
                                this.textBox2.Text = "";
                                dt2.Rows.Add(id);
                                dataGridView2.DataSource = dt2;
                            }
                            
                            else
                            {
                                if (check != "4" && check != "A")
                                {
                                    //Console.Beep(2000, 1000);
                                    MessageBox.Show("此錫膏現在狀態為:" + codemapping(check));
                                    this.textBox2.Text = "";
                                }

                            }

                        }
                        catch (Exception msg)
                        {
                            //Console.Beep(2000, 1000);
                            MessageBox.Show(msg.ToString());
                            this.textBox2.Text = "";
                        }
                    }
                }
            }

        }



        #region function  

        private string DateDiff(DateTime DateTime1, DateTime DateTime2)//完整時間比較(無使用)
        {
            string dateDiff = null;
            try
            {
                TimeSpan ts1 = new TimeSpan(DateTime1.Ticks);
                TimeSpan ts2 = new TimeSpan(DateTime2.Ticks);
                TimeSpan ts = ts1.Subtract(ts2).Duration();
                dateDiff = ts.Days.ToString() + "天"
                + ts.Hours.ToString() + "小時"
                + ts.Minutes.ToString() + "分鐘"
                + ts.Seconds.ToString() + "秒";
            }
            catch
            {

            }
            return dateDiff;
        }

        protected string codemapping(string type)   //狀態字串轉換
        {
            string str = "";

            switch (type)
            {
                default:
                    str = "異常";
                    break;
                case "1":
                    str = "入庫";
                    break;
                case "2":
                    str = "回溫";
                    break;
                case "3":
                    str = "攪拌";
                    break;
                case "4":
                    str = "使用";
                    break;
                case "5":
                    str = "結束";
                    break;
                case "6":
                    str = "報廢";
                    break;
                case "7":
                    str = "二次回庫";
                    break;
                case "8":
                    str = "二次回溫";
                    break;
                case "9":
                    str = "二次攪拌";
                    break;
                case "A":
                    str = "二次使用";
                    break;
            }
            return str;

        }

        private void date_Now(object sender, EventArgs e)
        {
            label8.Text = "Date:" + GetDate.Now();
        }//time

        #endregion
    }
}
