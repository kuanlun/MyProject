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
using System.Configuration;
using DAL;
using Models;


namespace SPM
{
    public partial class FrmSurplusSolder : Form
    {
        SurplusSolderServic ssServic = new SurplusSolderServic();
        public FrmSurplusSolder()
        {
            InitializeComponent();

        }
        

        private void SurplusSolder_Load(object sender, EventArgs e)
        {
            label6.Text = Class1.PName;//畫面傳值:receive
            label7.Text = Class1.PName1;//畫面傳值:receive
            
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();

            this.Visible = false;
            form1.Visible = true;
            this.Close();
        }

        //private string checkStatus(string strBar)//检查状态是否为4（使用）
        //{
        //    string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
        //    SqlConnection SqlCon = new SqlConnection(ConStr);
        //    SqlCon.Open();
        //    SqlCommand comm = new SqlCommand("select status from SolderPaste_Transaction where SolderID='" + strBar + "' order by TranTime desc", SqlCon);
        //    return comm.ExecuteScalar().ToString();


        //}
        private void txtBarcode_KeyDown(object sender, KeyEventArgs e)
        {
            
            if (e.KeyCode == Keys.Enter)
            {
                if (txtBarcode.Text.Length != 28)
                {
                    lbOut.Items.Insert(0, "錫膏條碼不正確，請重新輸入！！！");
                    txtBarcode.Text = null;
                    txtBarcode.Focus();
                    return;
                }
                //string strStatus = checkStatus(txtBarcode.Text.Trim());
                string strStatus = ssServic.GetStatusByBarcode(txtBarcode.Text.Trim()).ToString();
                if (strStatus != "4")
                {
                    lbOut.Items.Insert(0, "錫膏條碼狀態不正確，狀態需為'使用'狀態");
                    txtBarcode.Text = null;
                    txtBarcode.Focus();
                    return;
                }
                SurplusSolder objSS = new SurplusSolder()
                {
                    SolderID = txtBarcode.Text.Trim(),
                    Status="7",//锡䯧回库
                    TranTime=Convert.ToDateTime(GetDate.Now()),
                    UserID= Class1.PName,
                    ReMark="T"
                    
                };
                if (ssServic.InsertToDB(objSS) == 1)
                {
                    lbOut.Items.Insert(0, "錫膏【" + txtBarcode.Text.Trim()+"】回庫完成！！！");
                    txtBarcode.Text = null;
                    txtBarcode.Focus();
                    return;
                }
                
            }
        }

        
    }
}
