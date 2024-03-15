using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Globalization;

namespace SPM
{
    public partial class Form2 : Form
    {
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        int idcount=0;
        DateTime time;

        public Form2()
        {
            InitializeComponent();
            this.ControlBox = false;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            timer1.Start();
            label6.Text = Class1.PName;//畫面傳值:receive
            label7.Text = Class1.PName1;//畫面傳值:receive
            textBox1.KeyDown += textBox1_KeyDown;
            textBox2.KeyDown += textBox2_KeyDown;
            textBox3.KeyDown += textBox3_KeyDown;
            //入庫view
            //宣告欄位
            //dt.Columns.Add("ID", typeof(string));//12.26新增欄位序列號
            dt.Columns.Add("DVP", typeof(string));
            dt.Columns.Add("儲位", typeof(string));
            dt.Columns.Add("入庫時間", typeof(DateTime));
            dt.Columns.Add("status", typeof(string));
            dt.Columns.Add("UserID", typeof(string));

            //寫入值 for test
            //dt.Rows.Add("2019240275150000122310Lief01", "", null, "1", "G01110");
            //dt.Rows.Add("2019240275150000122310Lief02", "", null, "1", "G01110");
            //dt.Rows.Add("2019240275150000122310Lief03", "", null, "1", "G01110");

            //放入dataGridView1顯示      
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.DataSource = dt;
            dataGridView1.Columns["status"].Visible = false;
            dataGridView1.Columns["UserID"].Visible = false;

            //入庫取消view
            //宣告欄位
            dt2.Columns.Add("DVP",typeof(string));
            dt2.Columns.Add("儲位",typeof(string));
            dt2.Columns.Add("入庫時間",typeof(DateTime));
            dt2.Columns.Add("status", typeof(string));
            dt2.Columns.Add("UserID", typeof(string));

            //放入dataGridView2顯示
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView2.DataSource = dt2;
            dataGridView2.Columns["儲位"].Visible = false;
            dataGridView2.Columns["入庫時間"].Visible = false;
            dataGridView2.Columns["status"].Visible = false;
            dataGridView2.Columns["UserID"].Visible = false;

            textBox1.Select();
            label13.Text = "0";
            tabControl1.SelectedIndexChanged += tabControl1_SelectedIndexChanged;
        }

        private void button1_Click(object sender, EventArgs e)  //返回主選單
        {
            Form1 form1 = new Form1();
            this.Visible = false;
            form1.Visible = true;
            this.Close();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab.Name == "tabPage2")
            {
                textBox2.Select();
                dt.Clear();
                dataGridView1.DataSource = dt;
                textBox1.ReadOnly = false;
                button2.Enabled = true;
                label10.Text = "";
            }    
            if (tabControl1.SelectedTab.Name == "tabPage1")
            {
                textBox1.Select();
                dt2.Clear();
                dataGridView2.DataSource = dt2;
            }
                
        }

        private void button2_Click(object sender, EventArgs e)  //分配並執行入庫
        {

            if (GetParameter("000") == "Y")
            {
                if (Convert.ToInt32(label13.Text.Trim()) != Convert.ToInt32(GetParameter("001")))
                {
                    MessageBox.Show("本次鍚膏入庫筆數不等於:" + GetParameter("001") + "瓶?\r\n" + "請重新確認必須刷入瓶數!");
                    return;
                }
            }

            int CheckQty = CheckQtyinStock();
            if ((dataGridView1.RowCount - 1) > CheckQty)
            {
                MessageBox.Show("現有冰箱空儲位:" + CheckQty.ToString() + ",小於(不足)現在要入庫的總瓶數?\r\n請確認冰箱空儲位或執行出庫操作空出儲位空間。");
                return;
            }

            dt.DefaultView.Sort = "DVP ASC";
            dt = dt.DefaultView.ToTable();

            int dtRowCount = dt.Rows.Count;
            int no=0;
            string[] array = new string[dtRowCount]; 
            int loc_count = 0;    //可用儲位數
            string Loc_Y_max = "";    //最大可用儲位
            string Loc_Y_min = "";    //最小可用儲位
            string Loc_max_status = "";    //最大儲位狀態
            string Loc_min_status = "";    //最小儲位狀態
            string loc_status;

            //連接sql
            //string ConStr = "server=10.2.0.8;user id=sa;pwd=Mes123456;database=GTIMES";
            //string ConStr = "server=10.2.0.230;user id=sa;pwd=ks@12345;database=GTIMES";
            string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
            SqlConnection SqlCon = new SqlConnection(ConStr);

            SqlCon.Open();  //Open

            #region 取最大、最小可用儲位

            SqlDataReader dr = null;
            SqlCommand Cmd_item = new SqlCommand("select  MAX(CONVERT(int,Loc)) as L_max ,MIN(CONVERT(int,Loc)) as L_min ,Count(Loc) as Loc_count  from    SolderPaste_Location where IsEmpty='Y' ", SqlCon);
            dr = Cmd_item.ExecuteReader(); //執行並回傳

            if (dr.HasRows)
            {
                while (dr.Read())
                {
                    Loc_Y_max = dr["L_max"].ToString();
                    Loc_Y_min = dr["L_min"].ToString();
                    loc_count = Convert.ToInt32(dr["Loc_count"].ToString());
                }

            }

            Cmd_item.Cancel(); //關閉DataReader前先關閉sqlcommend
            dr.Close(); //關閉DataReader
            #endregion

            #region 判斷儲存狀態、取得儲位

            for (int i=0; i<dtRowCount; i++)
            {

                #region 取最大最小儲位狀態
                dr = null;
                Cmd_item = new SqlCommand("select isEmpty as L_max from SolderPaste_Location where Loc=(select MAX(CONVERT(int,Loc))  from  SolderPaste_Location) ", SqlCon);
                dr = Cmd_item.ExecuteReader(); //執行並回傳

                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        Loc_max_status = dr["L_max"].ToString();
                    }

                }

                Cmd_item.Cancel(); //關閉DataReader前先關閉sqlcommend
                dr.Close(); //關閉DataReader

                dr = null;
                Cmd_item = new SqlCommand("select  isEmpty as L_min from SolderPaste_Location where Loc=(select MIN(CONVERT(int,Loc))  from  SolderPaste_Location) ", SqlCon);
                dr = Cmd_item.ExecuteReader(); //執行並回傳

                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        Loc_min_status = dr["L_min"].ToString();
                    }

                }

                Cmd_item.Cancel(); //關閉DataReader前先關閉sqlcommend
                dr.Close(); //關閉DataReader

                #endregion

                loc_status = Loc_status_check(Loc_max_status, Loc_min_status);

                array[i] = GetLocNo(loc_status);

            }

            #endregion

            #region 寫入入庫儲位、入庫時間、UserID
            foreach (DataRow drow in dt.Rows)
            {

                if (drow[2].ToString() == "")
                {

                    //drow[1] = Int32.Parse(array[no]);  //儲位
                    drow[1] = array[no];  //儲位
                    drow[2] = Convert.ToDateTime(GetDate.Now());
                    drow[3] = "1";
                    drow[4] = label6.Text;   //UserID
                    no++;
                }

            }
            #endregion

            #region datatable資料寫入DB

            using (SqlBulkCopy SBC = new SqlBulkCopy(SqlCon))
            {
                SBC.DestinationTableName = "SolderPaste_Transaction"; //table name

                //mapping
                SBC.ColumnMappings.Add("DVP", "SolderID");
                SBC.ColumnMappings.Add("Status", "Status");
                SBC.ColumnMappings.Add("入庫時間", "TranTime");
                SBC.ColumnMappings.Add("UserID", "UserID");
                SBC.ColumnMappings.Add("儲位", "Loc");

                SBC.WriteToServer(dt); //寫入DB

                SBC.Close();

            }
            #endregion


            SqlCon.Close();  //Close  
            dataGridView1.DataSource = dt;

            textBox1.ReadOnly = true;
            button2.Enabled = false;
            BtnInConfirm.Enabled = false;
            idcount = 0;
            label13.Text = "0";
        }

        private void button3_Click(object sender, EventArgs e)  
        {

            string id2 = "";

            foreach (DataRow drow in dt2.Rows)
            {
                id2 += drow[0] + ",";    //取出所有DVP序號並寫入陣列
            }

            if (id2.Length > 0)
            {
                id2 = id2.Substring(0, id2.Length - 1);
                string[] list = id2.Split(',');

                for (int i = 0; i < list.Length; i++)
                {
                    DelSPinfo(list[i]);
                }

                //MessageBox.Show("執行完畢!");

                dt2.Clear();
                dataGridView2.DataSource = dt2;
            }
            else
            {
                //Console.Beep(2000, 1000);
                MessageBox.Show("請輸入欲取消入庫之DVP序號!");
                this.textBox2.Text = "";
            }
            

        }

        private void button4_Click(object sender, EventArgs e) //取消當前作業(沒用)
        {
            string check = "";
            string id = "";

            #region 依據狀態卡控
            foreach (DataRow drow in dt.Rows)
            {
                check += drow[1];   //儲位
                id += drow[0] + ",";

            }

            id = id.Substring(0, id.Length - 1);
            string[] list = id.Split(',');

            if (check.Length > 0)
            {
                for (int i=0;i<list.Length;i++)
                {
                    DelSPinfo(list[i]);
                }
                dt.Clear();
                dataGridView1.DataSource = dt;
                //MessageBox.Show("已取消錫膏入庫!");
            }
            else
            {
                dt.Clear();
                dataGridView1.DataSource = dt;
            }
            #endregion

        }

        private void textBox1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)//輸入DVP序號(入庫)
        {
            //
            // Detect the KeyEventArg's key enumerated constant.
            //
            if (e.KeyCode == Keys.Enter)
            {

                if (textBox1.TextLength == 28)
                {
                    string id = this.textBox1.Text;
                    string check = "";    //確認錫膏歷程(for 卡控)
                    int loc_count = 0;    //可用儲位數
                    string L_max = "";    //最大可用儲位
                    string L_min = "";    //最小可用儲位
                    string check_cancellog = "";
                    string str = "";
                    string detecode_str = "";
                    int check_datecode = 0;
                    int id_datecode = 0;
                    string dt_data_check = "";

                    #region 檢查是否在datatable已有入庫資料的情況下繼續作業，如有便先清空

                    string check_status = "";

                    foreach (DataRow drow in dt.Rows)
                    {
                        check_status += drow[1];   //儲位
                    }

                    if (check_status.Length > 0)
                    {
                        dt.Clear();
                        dataGridView1.DataSource = dt;
                    }

                    #endregion

                    #region SqlDataReader取值

                    //連接sql
                    //string ConStr = "server=10.2.0.8;user id=sa;pwd=Mes123456;database=GTIMES";
                    //string ConStr = "server=10.2.0.230;user id=sa;pwd=ks@12345;database=GTIMES";
                    string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
                    SqlConnection SqlCon = new SqlConnection(ConStr);

                    SqlCon.Open();  //open

                    //宣告DataReader
                    SqlDataReader dr = null;
                    SqlCommand Cmd_item = new SqlCommand("select  top 1 status as status from SolderPaste_Transaction where SolderID='" + id + "' order by TranTime  DESC ", SqlCon);
                    dr = Cmd_item.ExecuteReader(); //執行並回傳

                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            check = dr["status"].ToString();    //抓取ID最新狀況

                        }

                    }
                    Cmd_item.Cancel();
                    dr.Close();

                    //檢查最新Cancellog
                    dr = null;
                    Cmd_item = new SqlCommand("select top 1 AsIsStatus as check_cancellog from  SolderPaste_CancelLog  where  SolderID='" + id + "'  order by  TranTime  desc", SqlCon);
                    dr = Cmd_item.ExecuteReader(); //執行並回傳

                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            check_cancellog = dr["check_cancellog"].ToString();

                        }

                    }
                    Cmd_item.Cancel();
                    dr.Close();

                    //取昨日為止最大datecode
                    dr = null;
                    Cmd_item = new SqlCommand("select top 1 SUBSTRING(SolderID,0,7) as status FROM SolderPaste_Transaction  where TranTime <DateAdd(Day, DateDiff(Day, 0, getdate()), 0) order by status  desc", SqlCon);
                    dr = Cmd_item.ExecuteReader(); //執行並回傳

                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            detecode_str = dr["status"].ToString();

                        }

                    }

                    Cmd_item.Cancel(); //關閉DataReader前先關閉sqlcommend
                    dr.Close(); //關閉DataReader

                    #region 取最大、最小可用儲位

                    dr = null;
                    Cmd_item = new SqlCommand("select  MAX(CONVERT(int,Loc)) as L_max ,MIN(CONVERT(int,Loc)) as L_min ,Count(Loc) as Loc_count  from    SolderPaste_Location where IsEmpty='Y' ", SqlCon);
                    dr = Cmd_item.ExecuteReader(); //執行並回傳

                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            L_max = dr["L_max"].ToString();
                            L_min = dr["L_min"].ToString();
                            loc_count = Convert.ToInt32(dr["Loc_count"].ToString());
                        }

                    }

                    Cmd_item.Cancel(); //關閉DataReader前先關閉sqlcommend
                    dr.Close(); //關閉DataReader
                    #endregion

                    SqlCon.Close(); //close
                    #endregion

                    #region datecode 判斷

                    id_datecode = Convert.ToInt32(id.Substring(0, 6));
                    if (detecode_str == "")
                    {
                        check_datecode = 0;
                    }
                    else
                    {
                        check_datecode = Convert.ToInt32(detecode_str);
                    }


                    if (id_datecode < check_datecode)
                    {
                        str += "Datecode:" + id_datecode + "不合規定拒收,\n";
                        str += "目前冰箱最大datecode:" + check_datecode;
                    }

                    #endregion

                    #region 檢查是DVP是否重複刷入
                    foreach (DataRow drow in dt.Rows)
                    {

                        if (drow[0].ToString() == id)
                        {
                            dt_data_check = "錫膏重複!系統不讓刷入!";
                        }

                    }
                    #endregion

                    //卡控
                    if (check != null && check != "")
                    {
                        //Console.Beep(2000, 1000);
                        MessageBox.Show("此錫膏目前最新狀態為:" + codemapping(check) + "\n" + "無法進行入庫作業");
                        this.textBox1.Text = "";

                    }
                    else if (dt.Rows.Count > loc_count)
                    {
                        //Console.Beep(2000, 1000);
                        MessageBox.Show("可用庫存數不足!\n目前可用庫存數:" + loc_count + "\n" + "欲入庫錫膏數:" + dt.Rows.Count);
                        this.textBox1.Text = "";

                    }
                    else if (str != "")
                    {
                        //Console.Beep(2000, 1000);
                        MessageBox.Show(str);
                        this.textBox1.Text = "";

                    }
                    else if (dt_data_check != "")
                    {
                        //Console.Beep(2000, 1000);
                        MessageBox.Show(dt_data_check);
                        this.textBox1.Text = "";

                    }
                    else
                    {
                        if (id != "" && id != null)
                        {
                            try
                            {
                                DataRow dataRow = dt.NewRow();
                                //dataRow[0] = idcount++;
                                dataRow[0] = id;

                                dt.Rows.Add(dataRow);

                                //dt.DefaultView.Sort = "DVP asc"; //先將datatable排序

                                //放入dataGridView1顯示
                                dataGridView1.DataSource = dt;
                                idcount++;
                                label13.Text = idcount.ToString();
                            }
                            catch (Exception error)
                            {
                                //Console.Beep(2000, 1000);
                                MessageBox.Show(error.ToString());
                                this.textBox1.Text = "";
                            }

                        }

                        this.textBox1.Text = "";

                    }

                }
            }

        }

        private void textBox2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)//輸入DVP序號(入庫取消)
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
                    //寫入值

                    if (id.Length == 28)
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
                        SqlCommand Cmd_item = new SqlCommand("select  top 1 status as status from SolderPaste_Transaction where SolderID='" + id + "' order by TranTime  DESC", SqlCon);
                        dr = Cmd_item.ExecuteReader(); //執行並回傳

                        if (dr.HasRows)
                        {
                            while (dr.Read())
                            {
                                check = dr["status"].ToString();

                            }

                        }

                        Cmd_item.Cancel(); //關閉DataReader前先關閉sqlcommend
                        dr.Close(); //關閉DataReader
                        SqlCon.Close(); //close
                        #endregion

                        if (check == "1")
                        {
                            try
                            {
                                DataRow dataRow = dt2.NewRow();

                                dataRow[0] = id;
                                dt2.Rows.Add(dataRow);

                                //放入dataGridView2顯示
                                dataGridView2.DataSource = dt2;

                                //strat
                                //foreach (DataRow drow in dt2.Rows)
                                //{
                                //    id2 += drow[0] + ",";    //取出所有DVP序號並寫入陣列
                                //}

                                //if (id2.Length > 0)
                                //{
                                //    id2 = id2.Substring(0, id2.Length - 1);
                                //    string[] list = id2.Split(',');

                                //    for (int i = 0; i < list.Length; i++)
                                //    {
                                //        DelSPinfo(list[i]);
                                //    }

                                //    //MessageBox.Show("執行完畢!");

                                //    dt2.Clear();
                                //    //dataGridView2.DataSource = dt2;
                                //}
                                //else
                                //{
                                //    MessageBox.Show("請輸入欲取消入庫之DVP序號!");
                                //}
                                //end
                                DelSPinfo(id);

                                this.textBox2.Text = "";

                            }
                            catch (Exception error)
                            {
                                //Console.Beep(2000, 1000);
                                MessageBox.Show(error.ToString());
                                this.textBox2.Text = "";
                            }
                        }
                        else
                        {
                            //Console.Beep(2000, 1000);
                            MessageBox.Show("此錫膏現在狀態為:" + codemapping(check));
                            this.textBox2.Text = "";

                        }

                    }
                }
            }

        }

        private void textBox3_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)//放入冰箱
        {
            //
            // Detect the KeyEventArg's key enumerated constant.
            //
            if (e.KeyCode == Keys.Enter)
            {

                if (textBox3.TextLength == 28)
                {
                    string id = this.textBox3.Text;
                    string loc = "";

                    try
                    {
                        #region 確認錫膏的最新儲位

                        //連接sql
                        //string ConStr = "server=10.2.0.8;user id=sa;pwd=Mes123456;database=GTIMES";
                        //string ConStr = "server=10.2.0.230;user id=sa;pwd=ks@12345;database=GTIMES";
                        string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
                        SqlConnection SqlCon = new SqlConnection(ConStr);

                        SqlCon.Open();  //open

                        //宣告DataReader
                        SqlDataReader dr = null;
                        SqlCommand Cmd_item = new SqlCommand("select  top 1 Loc as Loc from SolderPaste_Transaction where SolderID='" + id + "' order by TranTime  DESC", SqlCon);
                        dr = Cmd_item.ExecuteReader(); //執行並回傳

                        if (dr.HasRows)
                        {
                            while (dr.Read())
                            {
                                loc = dr["Loc"].ToString();

                            }

                        }

                        Cmd_item.Cancel(); //關閉DataReader前先關閉sqlcommend
                        dr.Close(); //關閉DataReader
                        SqlCon.Close(); //close
                        #endregion

                        label10.Text = loc;
                        this.textBox3.Text = "";

                    }
                    catch (Exception er)
                    {
                        MessageBox.Show(er.ToString());
                        this.textBox3.Text = "";
                    }

                }
            }

        }

        //function
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
            }
            return str;

        }

        protected string Loc_status_check(string max, string min) //儲存狀態確認
        {
            string str = "";

            if(max=="Y" && min == "Y")
            {
                str = "1";
            }
            else if (max == "Y" && min == "N")
            {
                str = "2";
            }
            else if (max == "N" && min == "Y")
            {
                str = "3";
            }
            else
            {
                str = "4";
            }

            return str;
        }

        protected string GetLocNo(string str)   //取得儲位、鎖定儲位
        {
            string loc = "";
            string com;
            //連接sql
            //string ConStr = "server=10.2.0.8;user id=sa;pwd=Mes123456;database=GTIMES";
            //string ConStr = "server=10.2.0.230;user id=sa;pwd=ks@12345;database=GTIMES";
            string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
            SqlConnection SqlCon = new SqlConnection(ConStr);

            if (str == "1")
            {
                com = "select  MAX(CONVERT(int,Loc))+1 as loc from  SolderPaste_Location where IsEmpty='N'";
            }
            else
            {
                com = "select  MIN(CONVERT(int,Loc)) as loc from  SolderPaste_Location where IsEmpty='Y' ";
            }

            SqlCon.Open();  //Open

            #region 取儲位

            SqlDataReader dr = null;
            SqlCommand Cmd_item = new SqlCommand(com, SqlCon);
            dr = Cmd_item.ExecuteReader(); //執行並回傳

            if (dr.HasRows)
            {
                while (dr.Read())
                {
                    loc = dr["loc"].ToString();
                    if (loc=="")
                    {
                        loc = "1";
                    }
                }

            }

            Cmd_item.Cancel(); //關閉DataReader前先關閉sqlcommend
            dr.Close(); //關閉DataReader
            #endregion

            #region 鎖定儲位

            string InsStr = " update SolderPaste_Location set IsEmpty='N' where loc='"+loc+"'";
            SqlCommand Cmd = new SqlCommand(InsStr, SqlCon);
            Cmd.ExecuteNonQuery();

            #endregion


            SqlCon.Close();
            return loc;
        }

        protected void DelSPinfo(string str)    //刪除入庫紀錄、儲位解鎖
        {
            string loc = "";
            string com;
            //連接sql
            //string ConStr = "server=10.2.0.8;user id=sa;pwd=Mes123456;database=GTIMES";
            //string ConStr = "server=10.2.0.230;user id=sa;pwd=ks@12345;database=GTIMES";
            string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
            SqlConnection SqlCon = new SqlConnection(ConStr);

            com = "select top 1 loc from SolderPaste_Transaction where SolderID='"+str+ "' order by TranTime  desc";
            
            SqlCon.Open();  //Open

            #region 取對應儲位

            SqlDataReader dr = null;
            SqlCommand Cmd_item = new SqlCommand(com, SqlCon);
            dr = Cmd_item.ExecuteReader(); //執行並回傳

            if (dr.HasRows)
            {
                while (dr.Read())
                {
                    loc = dr["loc"].ToString();

                }

            }

            Cmd_item.Cancel(); //關閉DataReader前先關閉sqlcommend
            dr.Close(); //關閉DataReader
            #endregion

            /*
            //取服務器時間
            SqlDataReader dr2 = null;
            SqlCommand Cmd_item2 = new SqlCommand("select  getdate() as date", SqlCon);
            dr2 = Cmd_item2.ExecuteReader(); //執行並回傳

            if (dr2.HasRows)
            {
                while (dr2.Read())
                {
                    time = Convert.ToDateTime(dr2["date"].ToString());
                }

            }
            Cmd_item2.Cancel(); //關閉DataReader前先關閉sqlcommend
            dr2.Close(); //關閉DataReader
            */


            if (loc!=null && loc!="")   //確認有抓到值
            {
                #region 解鎖儲位

                string InsStr = " update SolderPaste_Location set IsEmpty='Y' where loc='" + loc + "'";
                SqlCommand Cmd = new SqlCommand(InsStr, SqlCon);
                Cmd.ExecuteNonQuery();
                Cmd.Cancel();

                #endregion

                #region 刪除交易紀錄

                InsStr = "delete SolderPaste_Transaction where SolderID='"+str+"' and TranTime=(SELECT top 1 TranTime FROM SolderPaste_Transaction where SolderID='"+str+"' order by  trantime DESC)";
                Cmd = new SqlCommand(InsStr, SqlCon);
                Cmd.ExecuteNonQuery();
                Cmd.Cancel();

                #endregion

                #region 寫入CancelLog紀錄

                DateTime dateTime = DateTime.Now;
                //string time = dateTime.ToString("yyyy-MM-dd hh:mm:ss.fff");
                //InsStr = "insert into SolderPaste_CancelLog values('" + str + "','1',CONVERT(DATETIME,'" + time + "'))";
                InsStr = "insert into SolderPaste_CancelLog values('" + str + "','1',CONVERT(DATETIME,'" + GetDate.Now() + "'))";
                Cmd = new SqlCommand(InsStr, SqlCon);
                Cmd.ExecuteNonQuery();
                Cmd.Cancel();

                #endregion
            }

            SqlCon.Close();
        }

        private void date_Now(object sender, EventArgs e)
        {
            label8.Text = "Date:" + GetDate.Now();
        }//time

        private void BtnInConfirm_Click(object sender, EventArgs e)
        {
            int RowCount = dataGridView1.RowCount;
            if (RowCount == 1) {
                MessageBox.Show("鍚膏入庫筆數為0,請正確刷入條碼後再執行《入庫確認》!");
                return;
            }

            if (GetParameter("000") == "Y")
            {
                if (Convert.ToInt32(label13.Text.Trim()) != Convert.ToInt32(GetParameter("001")))
                {
                    MessageBox.Show("本次鍚膏入庫筆數不等於:" + GetParameter("001") + "瓶?\r\n" + "請重新確認必須刷入瓶數!");
                    return;
                }
            }

            int CheckQty = CheckQtyinStock();
            if ((RowCount-1) > CheckQty)
            {
                MessageBox.Show("現有冰箱空儲位:"+ CheckQty.ToString()+ "小於(不足)現在要入庫的總瓶數?\r\n" + "請確認冰箱空儲位或執行出庫操作空出儲位空間。");
                return;
            }

            //MessageBox.Show("請確認本次入庫鍚膏瓶數:"+label13.Text.Trim()+"\r\n,是否正確?");
            MessageBox.Show("請再次確認本次入庫鍚膏瓶數:"+ label13.Text.Trim() + ",是否正確?\r\n是,請執行下一步驟2.分配儲位!");
        }

        public static int CheckQtyinStock() //確認本次入庫數是否大於空儲位總數
        {
            string str = "";

            //連接sql
            string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
            SqlConnection SqlCon = new SqlConnection(ConStr);
            SqlCon.Open();

            //宣告DataReader
            SqlDataReader dr = null;
            SqlCommand Cmd = new SqlCommand("select max=count(*) from SolderPaste_Location(nolock) where IsEmpty = 'Y' ", SqlCon);
            dr = Cmd.ExecuteReader(); //執行並回傳

            if (dr.HasRows)
            {
                while (dr.Read())
                {
                    str = dr["max"].ToString();

                }
            }

            Cmd.Cancel();
            dr.Close();
            SqlCon.Close();
            SqlCon.Dispose();
            return Convert.ToInt32(str.Trim());
        }

        public static string GetParameter(string id) //取得系統參數表中ID=001(錫膏入庫上限數)
        {
            string str = "";

            //連接sql
            string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
            SqlConnection SqlCon = new SqlConnection(ConStr);
            SqlCon.Open();

            //宣告DataReader
            SqlDataReader dr = null;
            SqlCommand Cmd = new SqlCommand("SELECT  Params as Param from SolderPaste_Parameter(nolock) where ID="+"'"+id.Trim()+"' ", SqlCon);
            dr = Cmd.ExecuteReader(); //執行並回傳

            if (dr.HasRows)
            {
                while (dr.Read())
                {
                    str = dr["Param"].ToString();

                }
            }

            Cmd.Cancel();
            dr.Close();
            SqlCon.Close();
            SqlCon.Dispose();

            //str = Convert.ToDateTime(str).ToString("yyyy-MM-dd HH:mm:ss");
            return str.ToString().Trim();
        }

    }
}
