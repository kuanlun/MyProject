using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SPM
{
    class Class1  //畫面傳值:receive
    {
        static string name;
        static string name1;

        public static string PName
        {
            get
            {
                return name;
            }
            set
            {
                name = value;
            }
        }
        public static string PName1
        {
            get
            {
                return name1;
            }
            set
            {
                name1 = value;
            }
        }
    }

    class GetDate
    {
        public static string Now() //取資料庫系統時間
        {
            string str = "";

            //連接sql
            string ConStr = ConfigurationManager.ConnectionStrings["SQLDB"].ConnectionString;
            SqlConnection SqlCon = new SqlConnection(ConStr);
            SqlCon.Open();

            //宣告DataReader
            SqlDataReader dr = null;
            SqlCommand Cmd = new SqlCommand("SELECT  getdate() as get_date", SqlCon);
            dr = Cmd.ExecuteReader(); //執行並回傳

            if (dr.HasRows)
            {
                while (dr.Read())
                {
                    str = dr["get_date"].ToString();

                }
            }

            Cmd.Cancel();
            dr.Close();
            SqlCon.Close();
            SqlCon.Dispose();

            str = Convert.ToDateTime(str).ToString("yyyy-MM-dd HH:mm:ss");
            return str;
        }
    }

    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Login());
            
        }
    }
}
