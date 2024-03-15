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
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

namespace SPM
{
    public partial class QueryHistory : Form
    {

        DataTable dt = new DataTable();

        public QueryHistory()
        {
            InitializeComponent();
        }

        private void QueryHistory_Load(object sender, EventArgs e)
        {
            label6.Text = Class1.PName;//畫面傳值:receive
            label7.Text = Class1.PName1;//畫面傳值:receive
            for (int i = 0; i <= (checkedListBox1.Items.Count - 1); i++)
            {
                checkedListBox1.SetItemChecked(i, true);
                //checkedListBox1.SetItemCheckState
            }
            textBox1.Focus();
            dt.Columns.Add("錫膏代碼", typeof(string));
            dt.Columns.Add("狀態", typeof(string));
            dt.Columns.Add("進出時間點", typeof(string));
            dt.Columns.Add("儲位", typeof(string));

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();

            this.Visible = false;
            form1.Visible = true;
            this.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            string s, SqlStr = "", SqlStr1 = "", SqlStr2 = "", SqlStr3 = "", SqlStr4 = "", SqlStr5 = "";

            if (checkedListBox1.CheckedItems.Count == 0)
            {
                //Console.Beep(2000, 1000); //2000HZ,1 sec
                MessageBox.Show("請確實勾選錫膏狀態 !", "錯誤提示");

            }
            else
            {
                int i;
                //s = "Checked items:\n";
                //DB connection
                //string constr = "Data Source = 10.2.0.8; User Id = sa; Password = Mes123456; Initial Catalog = GTIMES; Max Pool Size = 300";
                //string constr = "server=10.2.0.230;user id=sa;pwd=Mes123456;database=GTIMES";
                string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
                SqlConnection SqlCon2 = new SqlConnection(ConStr); SqlCon2.Open();
                dt.Clear();
                /*
                DataTable dt = new DataTable();
                dt.Columns.Add("錫膏代碼", typeof(string));
                dt.Columns.Add("狀態", typeof(string));
                dt.Columns.Add("進出時間點", typeof(string));
                dt.Columns.Add("儲位", typeof(string));
                //dt.Columns.Add("攪拌時間點", typeof(string));
                //dt.Columns.Add("使用時間點", typeof(string));
                // dt.Columns.Add("結束時間點", typeof(string));
                //dt.Columns.Add("操作員", typeof(string));
                */
                //get time condition
                //string theDate1 = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                //string theDate2 = dateTimePicker2.Value.ToString("yyyy-MM-dd");

                if ((checkBox1.Checked == true)) SqlStr1 = " and (a.TranTime >='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' and a.TranTime <='" + dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm:ss") + "') ";
                else SqlStr1 = "";

                //SqlStr2 = "";
                //取得已選取的listbox框
                string[] arrStr = new string[10];
                string objStr = "";
                for (i = 0; i <= (checkedListBox1.Items.Count)-1; i++)
                {

                    if (checkedListBox1.GetItemChecked(i))
                    {
                        //this.checkedListBox1.SetSelected(i, true);

                       // string objCheck = checkedListBox1.CheckedItems[i].ToString();
                        string objCheck = checkedListBox1.GetItemText(checkedListBox1.Items[i]);
                        switch (objCheck)
                        {
                            case "入庫":
                                objStr = "1";
                                break;
                            case "回溫":
                                objStr = "2";
                                break;
                            case "攪拌":
                                objStr = "3";
                                break;
                            case "使用":
                                objStr = "4";
                                break;
                            case "結束":
                                objStr = "5";
                                break;
                            case "報廢":
                                objStr = "6";
                                break;
                            case "二次回庫":
                                objStr = "7";
                                break;
                            case "二次回溫":
                                objStr = "8";
                                break;
                            case "二次攪拌":
                                objStr = "9";
                                break;
                            case "二次使用":
                                objStr = "A";
                                break;


                        }
                        arrStr[i] = objStr;
                    }

                    //int ii = 0;
                    //ii = i + 1;
                    //SqlStr2 += ii.ToString();

                }

                SqlStr = @"Select a.solderid,case when a.status = '1' then '入庫'  when a.status = '2' then '回溫'  when a.status = '3' then '攪拌'  when a.status = '4' then '使用'  when a.status = '5' then '結束' when a.status='7' then '二次回庫' when a.status='8' then '二次回溫' when a.status='9' then '二次攪拌' when a.status='A' then '二次使用' end as status," +
                             "a.TranTime,a.loc from  SolderPaste_Transaction a with(nolock) where 1 = 1 " + SqlStr1 + " and a.SolderID like '%" + textBox1.Text.Trim() + "%' and a.status in ('{0}') ";
                //拼接in sql語法
                //SqlStr = string.Format(SqlStr, string.Join("','", SqlStr2.ToArray()));
                SqlStr = string.Format(SqlStr, string.Join("','", arrStr));
                if (checkBox1.Checked==true)
                {
                    SqlStr += " order by a.SolderID, a.status ";
                }
                else
                {
                    SqlStr += " order by a.TranTime,a.status ";
                }
               
                /*
                SqlStr = "SELECT distinct a.SolderID,a.Status as Status1,a.TranTime as TranTime1,a.UserID,b.Status as Status2,b.TranTime as TranTime2," +
                    "c.Status as Status3,c.TranTime as TranTime3,d.Status as Status4,d.TranTime as TranTime4," +
                    "e.Status as Status5,e.TranTime as TranTime5 FROM [GTIMES].[dbo].[SolderPaste_Transaction]  as a" +
                    " LEFT join  SolderPaste_Transaction  as b   on   a.SolderID=b.Solderid   and  a.status=1 and  b.status=2 " + SqlStr2.Trim()+
                    " LEFT join  SolderPaste_Transaction  as c   on   a.SolderID=c.Solderid     and  a.status=1 and  c.status=3" + SqlStr3.Trim() +
                    " LEFT join  SolderPaste_Transaction  as d   on   a.SolderID=d.Solderid        and  a.status=1 and  d.status=4" + SqlStr4.Trim() +
                    " LEFT join  SolderPaste_Transaction  as e   on   a.SolderID=e.Solderid     and  a.status=1 and  e.status=5 " + SqlStr5.Trim() +
                    " Where a.Status = 1 and a.SolderID like '%" + textBox1.Text.Trim() + "%' "+ SqlStr1.Trim()+ " order by SolderID";
                */
                //int ct = 0;
                //MessageBox.Show(s);   
                //SqlStr = SqlStr + "where a.Status='1' order by SolderID";

                DataSet ds = new DataSet();
                SqlDataAdapter Sda = new SqlDataAdapter(SqlStr, SqlCon2);
                Sda.Fill(ds);
                //dataGridView1.DataSource = ds;                

                //欄寬自動調整
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dataGridView1.DataSource = dt;
                string solderid;
                Boolean exist = false;

                for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                {
                    dt.Rows.Add(ds.Tables[0].Rows[j]["SolderID"].ToString(), ds.Tables[0].Rows[j]["status"].ToString(), ds.Tables[0].Rows[j]["TranTime"].ToString(), ds.Tables[0].Rows[j]["Loc"].ToString());
                }
                #region MyRegion


                //dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                /*
                for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                {
                    exist = false;
                    for (i = 0; i <= (checkedListBox1.Items.Count - 1); i++)
                    {
                        if (checkedListBox1.GetItemChecked(i))
                        {
                            //ct++;
                            switch (i)
                            {
                                default:
                                    break;
                                case 0: //入庫
                                    //SqlStr = "select * from SolderPaste_Transaction where Status in ('" + (i + 1).ToString() + "'";
                                    if (ds.Tables[0].Rows[j]["TranTime1"].ToString() != "")
                                    {
                                        for (int rows = 0; rows < dataGridView1.Rows.Count - 1; rows++)
                                        {
                                            solderid = dataGridView1.Rows[rows].Cells[0].Value.ToString();
                                            if (ds.Tables[0].Rows[j]["SolderID"].ToString() == solderid)
                                            {
                                                dataGridView1.Rows[rows].Cells[1].Value = ds.Tables[0].Rows[j]["TranTime1"].ToString();
                                                exist = true;
                                            }                                          
                                        } //for
                                        if (exist == false)
                                           dt.Rows.Add(ds.Tables[0].Rows[j]["SolderID"].ToString(), ds.Tables[0].Rows[j]["TranTime1"].ToString(), "", "", "", "");
                                    }
                                        break;
                                case 1: //回溫
                                    //SqlStr = SqlStr+",'" + (i + 1).ToString() + "'";
                                    if (ds.Tables[0].Rows[j]["TranTime2"].ToString() != "")
                                    {
                                        for (int rows = 0; rows < dataGridView1.Rows.Count - 1; rows++)
                                        {
                                            solderid = dataGridView1.Rows[rows].Cells[0].Value.ToString();
                                            if (ds.Tables[0].Rows[j]["SolderID"].ToString() == solderid)
                                            {
                                                dataGridView1.Rows[rows].Cells[2].Value = ds.Tables[0].Rows[j]["TranTime2"].ToString();
                                                exist = true;
                                            }
                                        } //for
                                        if (exist == false)
                                            dt.Rows.Add(ds.Tables[0].Rows[j]["SolderID"].ToString(), "", ds.Tables[0].Rows[j]["TranTime2"].ToString(), "", "", "");
                                    }

                                    break;
                                case 2: //攪拌
                                    //SqlStr = SqlStr + ",'" + (i + 1).ToString() + "'";
                                    if (ds.Tables[0].Rows[j]["TranTime3"].ToString() != "")
                                    {
                                        for (int rows = 0; rows < dataGridView1.Rows.Count - 1; rows++)
                                        {
                                            solderid = dataGridView1.Rows[rows].Cells[0].Value.ToString();
                                            if (ds.Tables[0].Rows[j]["SolderID"].ToString() == solderid)
                                            {
                                                dataGridView1.Rows[rows].Cells[3].Value = ds.Tables[0].Rows[j]["TranTime3"].ToString();
                                                exist = true;
                                            }
                                        } //for
                                        if (exist == false)
                                            dt.Rows.Add(ds.Tables[0].Rows[j]["SolderID"].ToString(), "", "", ds.Tables[0].Rows[j]["TranTime3"].ToString(), "", "");
                                    }

                                    break;
                                case 3: //使用
                                    //SqlStr = SqlStr + ",'" + (i + 1).ToString() + "'";
                                    if (ds.Tables[0].Rows[j]["TranTime4"].ToString() != "")
                                    {
                                        for (int rows = 0; rows < dataGridView1.Rows.Count - 1; rows++)
                                        {
                                            solderid = dataGridView1.Rows[rows].Cells[0].Value.ToString();
                                            if (ds.Tables[0].Rows[j]["SolderID"].ToString() == solderid)
                                            {
                                                dataGridView1.Rows[rows].Cells[4].Value = ds.Tables[0].Rows[j]["TranTime4"].ToString();
                                                exist = true;
                                            }
                                        } //for
                                        if (exist == false)
                                            dt.Rows.Add(ds.Tables[0].Rows[j]["SolderID"].ToString(), "", "", "", ds.Tables[0].Rows[j]["TranTime4"].ToString(), "");
                                    }                                                                            
                                    break;
                                case 4: //結束
                                    //SqlStr = SqlStr + ",'" + (i + 1).ToString() + "'";
                                    if (ds.Tables[0].Rows[j]["TranTime5"].ToString() != "")
                                    {
                                        for (int rows = 0; rows < dataGridView1.Rows.Count - 1; rows++)
                                        {
                                            solderid = dataGridView1.Rows[rows].Cells[0].Value.ToString();
                                            if (ds.Tables[0].Rows[j]["SolderID"].ToString() == solderid)
                                            {
                                                dataGridView1.Rows[rows].Cells[5].Value = ds.Tables[0].Rows[j]["TranTime5"].ToString();
                                                exist = true;
                                            }
                                        } //for
                                        if (exist == false)
                                            dt.Rows.Add(ds.Tables[0].Rows[j]["SolderID"].ToString(), "", "", "", "", ds.Tables[0].Rows[j]["TranTime5"].ToString());
                                    }

                                    break;
                            }
                            //s = s + "Item " + (i + 1).ToString() + " = " + checkedListBox1.Items[i].ToString() + "\n";
                        }
                    }                    
                }
                dataGridView1.DataSource = dt;               
                */
                #endregion
                SqlCon2.Close();
                SqlCon2.Dispose();
                int gdct = 0;
                gdct = dataGridView1.Rows.Count - 1;
                label4.Text = gdct.ToString();
                //textBox1.Text = "";

            } //else
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            for (int ii = 0; ii <= (checkedListBox1.Items.Count - 1); ii++)
            {
                checkedListBox1.SetItemChecked(ii, true);
                //checkedListBox1.SetItemCheckState
            }
            textBox1.Focus();
            checkBox1.Checked = false;
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            /*
            if (textBox1.Text.Length < 28)
            {
                MessageBox.Show("請輸入正確的DVP序號!");
                textBox1.Focus();
                return;
            }
            */
        }

        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                dateTimePicker1.Enabled = true;
                dateTimePicker2.Enabled = true;
            }
            else
            {
                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExportToExcel(dt);
        }

        public static void ExportToExcel(DataTable dt)
        {

            string dir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            if (dt.Rows.Count <= 0)
            {
                MessageBox.Show("空数据,无法导出Excel！\n\r " + "请正确操作!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }
            //IWorkbook workbook;

            HSSFWorkbook workbook = new HSSFWorkbook();
            if (workbook == null) { return; }
            ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet(dt.TableName);

            //表头  
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }

            //数据  
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    // row1["Status1"].ToString() 

                    ICell cell = row1.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            //转为字节数组  
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();
            DateTime dtTIME = Convert.ToDateTime(GetDate.Now());
            string fileName = @dir + @"\" + "锡膏_" + string.Format("{0:yyyyMMddHHmmss}", dtTIME) + ".xls";

            try
            {
                //保存为Excel文件  
                using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(buf, 0, buf.Length);
                    fs.Flush();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("导出异常：" + ex.Message, "导出异常", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                MessageBox.Show("导出Excel成功！!\n\r " + fileName, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}

