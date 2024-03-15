using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace SPM
{
    public partial class QueryStock : Form
    {
        DataTable dt = new DataTable();
        public QueryStock()
        {
            InitializeComponent();
        }

        private void QueryStock_Load(object sender, EventArgs e)
        {
            int i;
            label6.Text = Class1.PName;//畫面傳值:receive
            label7.Text = Class1.PName1;//畫面傳值:receive
            //DB connection
            //string constr = "Data Source = 10.2.0.8; User Id = sa; Password = Mes123456; Initial Catalog = GTIMES; Max Pool Size = 300";
            //string constr = "server=10.2.0.230;user id=sa;pwd=ks@12345;database=GTIMES";
            string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
            SqlConnection SqlCon2 = new SqlConnection(ConStr);
            SqlCon2.Open();
            dt.Columns.Add("錫膏代碼", typeof(string));
            dt.Columns.Add("入庫時間點", typeof(string));
            dt.Columns.Add("儲位", typeof(string));

            string SqlStr = "SELECT * from SolderPaste_Transaction where Status='1' " +
                "and SolderID not in (select SolderID from SolderPaste_Transaction where Status = '2' or Status = '3' or Status = '4' or Status = '5' or Status = '6') order by trantime ";

            DataSet ds = new DataSet();
            SqlDataAdapter Sda = new SqlDataAdapter(SqlStr, SqlCon2);
            Sda.Fill(ds);
            //dataGridView1.DataSource = ds;                

            //欄寬自動調整
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.DataSource = dt;
            for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
            {
                dt.Rows.Add(ds.Tables[0].Rows[j]["SolderID"].ToString(), ds.Tables[0].Rows[j]["TranTime"].ToString(), ds.Tables[0].Rows[j]["Loc"].ToString());
            }
            SqlCon2.Close();
            SqlCon2.Dispose();
            int gdct = 0;
            gdct = dataGridView1.Rows.Count - 1;
            label4.Text = gdct.ToString();
        } //sub

        private void button2_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();

            this.Visible = false;
            form1.Visible = true;
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            TableToExcel(dt);
        }

        public static void TableToExcel(DataTable dt)
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
