using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Net;
using System.Text;
using System.Windows.Forms;
//using System.Diagnostics;
using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace Salary4Zhen
{
    public partial class MainForm : Form
    {
        public static string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}\\Salary.accdb";
        public static string strConnExcel = "Provider=Microsoft.Ace.OleDb.12.0;Data Source = {0};Extended Properties='Excel 12.0;HDR=Yes'";
        public static string strFilter = "Excel File(*.xls;*.xlsx)|*.xls;*.xlsx";
        public static string strWeb = @"http://bluelaser.cn:88/bluelaser/Yaka.txt";
        public bool bValid = true;
        public string strMonth = "";

        //for 整点提醒
        public string strUserName = "Yaka";
        public static string strMessage = "{0}用时{1:F4}秒,总计执行{2}条记录,请务必休息{3}分钟后再工作！";
        public static string strMessageHours = "，请站起来上个厕所，喝个开水，放松眼睛，休息5分钟后再工作！";
        public static string strMessageLunch = "，午饭时间到了，请按时吃饭，5分钟后自动锁屏！";
        public static string strMessageOffduty = "，下班时间到了，请放下工作，回家吃饭，5分钟后自动关机！";
        
        //for import
        public static string strSQLDel0 = "DELETE * FROM [{0}工资] ";
        public static string strSQLImport = "INSERT INTO [{0}工资] SELECT * FROM [excel 12.0;database={1}].[sheet1$]";
        //for export
        public static string strSQLExport = "SELECT * INTO [excel 12.0;database={0}].[{1}] FROM [{1}工资]";

        //for 扣回汇总
        public static string strSQL0 = "DELETE * FROM 扣回汇总 WHERE 所属月份='{0}'";
        public static string strSQL1 = @"INSERT INTO 扣回汇总 ( 所属月份, 骑手ID, 骑手信息, 所属团队, 在职情况, 房租, 电瓶车, 物资, 充电 )
                                     SELECT '{0}', 骑手ID, 骑手信息, 所属团队, 在职情况, 房租, 电瓶车, 物资, 充电 
                                     FROM {0}工资 WHERE 应发工资 >=0 AND(房租<0 OR 电瓶车< 0 OR  物资< 0 OR 充电< 0)";
        public static string strSQL2 = @"INSERT INTO 扣回汇总 ( 所属月份, 骑手ID, 骑手信息, 所属团队, 在职情况, 房租, 电瓶车, 物资, 充电 )
                                     SELECT '{0}', 骑手ID, 骑手信息, 所属团队, 在职情况, 房租, 电瓶车, 物资, 充电 
                                     FROM {0}工资 WHERE 应发工资< 0 AND (房租<0 OR 电瓶车< 0 OR  物资< 0 OR 充电< 0)  
                                     AND 骑手ID IN (SELECT 骑手ID FROM {1}工资 WHERE 上月遗留< 0 AND (应发工资 - 房租 - 电瓶车- 物资- 充电 - 预支)>= 0)";
        public static string strSQL3 = @"INSERT INTO 扣回汇总 ( 所属月份, 骑手ID, 骑手信息, 所属团队, 在职情况, 房租, 电瓶车, 物资, 充电, 部分扣回 ) 
                                     SELECT '{0}', 骑手ID, 骑手信息, 所属团队, 在职情况, 房租, 电瓶车, 物资, 充电, 应发工资 - 房租 - 电瓶车- 物资- 充电 - 预支 
                                     FROM {0}工资 WHERE 应发工资< 0 AND (房租<0 OR 电瓶车< 0 OR  物资< 0 OR 充电< 0) 
                                     AND 应发工资 - 房租 - 电瓶车- 物资- 充电 - 预支>0 AND 骑手ID NOT IN(SELECT 骑手ID FROM {1}工资)";
        public static string strSQL4 = @"INSERT INTO 扣回汇总 ( 所属月份, 骑手ID, 骑手信息, 所属团队, 在职情况, 房租, 电瓶车, 物资, 充电, 部分扣回 ) 
                                     SELECT '{0}', [{0}工资].骑手Id, 骑手信息, 所属团队, 在职情况, 房租, 电瓶车, 物资, 充电, 部分扣回 
                                     FROM {0}工资 INNER JOIN(SELECT 骑手ID, (应发工资 - 房租 - 电瓶车- 物资- 充电 - 预支 - 上月遗留) AS 部分扣回 FROM {1}工资 WHERE 上月遗留< 0 AND 应发工资 - 房租 - 电瓶车- 物资- 充电 - 预支< 0)  AS Nextview  
                                     ON [{0}工资].骑手id = Nextview.骑手id WHERE Nextview.部分扣回 > 0";
        public static string strSQLSelect = "SELECT * FROM 扣回汇总 WHERE 部分扣回 > 0 AND 所属月份 = '{0}'";
        public static string strSQLInto0 = "SELECT * INTO [excel 12.0;database={0}].[{1}] FROM 扣回汇总 WHERE 所属月份 = '{1}'";

        //for 薪资报备
        public static string strSQLDel = "DELETE * FROM 薪资报备";
        public static string strSQLInto = @"INSERT INTO 薪资报备 SELECT * FROM [excel 12.0;database={0}].[sheet1$]";
        public static string strSQL5 = @"UPDATE 薪资报备 INNER JOIN [{0}工资] ON(薪资报备.团队名称 = [{0}工资].所属团队) AND(薪资报备.骑手ID = [{0}工资].骑手Id) 
                                    SET 装备物资 = 物资, 站点充电 = 充电, 住宿费用 = 房租, 薪资报备.保险 = [{0}工资].保险, 薪资报备.个税 = [{0}工资].个税,骑手预支借款 = -预支,骑手预支还款 = 0, 实发薪资 = 实发工资-[{0}工资].保险-房租-物资-充电-预支";
        //public static string strSQL51 = @"UPDATE 薪资报备 INNER JOIN (SELECT * FROM {0}工资 WHERE 骑手ID not in (SELECT 骑手ID FROM {0}报备 WHERE 实发薪资 IS NOT NULL)) AS 工资View
        //                           ON [薪资报备].骑手ID = 工资View.骑手Id SET 装备物资 = 物资, 站点充电 = 充电, 住宿费用 = 房租, [{0}报备].保险 = 工资View.保险, [{0}报备].个税 = 工资View.个税,骑手预支借款 = -预支, 骑手预支还款 = 0,实发薪资 = 实发工资-工资View.保险-房租-物资-充电-预支";
        public static string strSQL6 = "UPDATE 薪资报备 SET 实发薪资 = 0 WHERE 实发薪资 < 0";
        public static string strSQL7 = "UPDATE 薪资报备 SET [其他加项(金额)] = 实发薪资-(底薪+提成+价格奖励+距离奖励+时段奖励+天气奖励+好评奖励+冲单奖励+差评扣款+投诉扣款+违规扣款+商户索赔+装备物资+站点充电+住宿费用+保险+个税+骑手预支借款) WHERE 实发薪资 IS NOT NULL ";
        public static string strSQL8 = "UPDATE 薪资报备 SET [其他加项(金额)] = 0,[其他减项(金额)] = [其他加项(金额)] WHERE [其他加项(金额)] < 0 ";
        public static string strSQLInto2 = "SELECT * INTO [excel 12.0;database={0}].薪资报备 FROM 薪资报备";
        
        //for 工资计算
        public static string strSQL9 = "UPDATE [工资表$] INNER JOIN [成立差评$] ON [成立差评$].骑手信息 = [工资表$].骑手信息 AND [成立差评$].团队 = [工资表$].团队名称 SET 成立差评 = 汇总";
        [DllImport("user32.dll")]
        public static extern void LockWorkStation();
        public MainForm()
        {
            InitializeComponent();
            listMonth.SetSelected(DateTime.Now.Month-2, true);
        }

        private bool CheckYaka()
        {
            try
            {
                HttpWebRequest LoginReq = (HttpWebRequest)WebRequest.Create(strWeb);
                HttpWebResponse LoginRes = (HttpWebResponse)LoginReq.GetResponse();
                Stream receiveStream = LoginRes.GetResponseStream();
                StreamReader readStream = new StreamReader(receiveStream);
                bValid = (readStream.ReadToEnd() == "Yaka");
                readStream.Close();
                LoginRes.Close();
            }
            catch (WebException ex)
            {
                MessageBox.Show(ex.Message);
            }
            return bValid;
        }
        private void btnAnalysis_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = strFilter;
            if (ofd.ShowDialog() == DialogResult.OK && bValid)
            {
                OleDbConnection conn = new OleDbConnection(string.Format(strConn, Application.StartupPath));

                string strNextMonth = listMonth.Items[listMonth.SelectedIndex + 1].ToString();
                if (radioHuzhou.Checked)
                    strNextMonth = 'H' + strNextMonth;
                try
                {
                    DateTime timeStart = DateTime.Now;
                    conn.Open();
                    OleDbCommand commandInsert = new OleDbCommand(string.Format(strSQL0, strMonth), conn);
                    int intLow = commandInsert.ExecuteNonQuery();
                    commandInsert.CommandText = string.Format(strSQL1, strMonth, strNextMonth);
                    intLow = commandInsert.ExecuteNonQuery();
                    commandInsert.CommandText = string.Format(strSQL2, strMonth, strNextMonth);
                    intLow += commandInsert.ExecuteNonQuery();
                    commandInsert.CommandText = string.Format(strSQL3, strMonth, strNextMonth);
                    intLow += commandInsert.ExecuteNonQuery();
                    commandInsert.CommandText = string.Format(strSQL4, strMonth, strNextMonth);
                    intLow += commandInsert.ExecuteNonQuery();

                    OleDbCommand command = new OleDbCommand(string.Format(strSQLSelect, strMonth), conn);
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        StringBuilder strSQL = new StringBuilder("UPDATE 扣回汇总 SET ");
                        int intpart = (int)reader["部分扣回"];
                        if (intpart <= -(int)reader["房租"])
                            strSQL.Append("房租=-" + intpart.ToString() + ",电瓶车=0,物资=0,充电=0");
                        else
                        {
                            intpart += (int)reader["房租"];
                            if (intpart <= -(int)reader["电瓶车"])
                                strSQL.Append("电瓶车=-" + intpart.ToString() + ",物资=0,充电=0");
                            else
                            {
                                intpart += (int)reader["电瓶车"];
                                if (intpart <= -(int)reader["物资"])
                                    strSQL.Append("物资=-" + intpart.ToString() + ",充电=0");
                                else
                                {
                                    intpart += (int)reader["物资"];
                                    if (intpart <= -(int)reader["充电"])
                                        strSQL.Append("充电=-" + intpart.ToString());
                                    else
                                        continue;
                                }
                            }
                        }
                        strSQL.Append(" WHERE 所属月份 = '" + strMonth + "' AND 骑手Id=" + reader["骑手Id"].ToString());
                        OleDbCommand command2 = new OleDbCommand(strSQL.ToString(), conn);
                        intLow += command2.ExecuteNonQuery();
                    }
                    reader.Close();

                    commandInsert.CommandText = string.Format(strSQLInto0, ofd.FileName,strMonth);
                    commandInsert.ExecuteNonQuery();
                    MessageBox.Show(string.Format(strMessage, strMonth + btnAnalysis.Text, (DateTime.Now - timeStart).TotalSeconds, intLow, intLow / 10), "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (conn.State == ConnectionState.Open)
                        conn.Close();
                }
            }
        }

        private void btnReport_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = strFilter;
            if (ofd.ShowDialog() == DialogResult.OK && bValid)
            {
                OleDbConnection connAccess = new OleDbConnection(string.Format(strConn, Application.StartupPath));
                try
                {
                    DateTime timeStart = DateTime.Now;
                    connAccess.Open();
                    OleDbCommand commandInsert = new OleDbCommand(strSQLDel, connAccess);
                    int intLow = commandInsert.ExecuteNonQuery();
                    commandInsert.CommandText = string.Format(strSQLInto, ofd.FileName);
                    intLow = commandInsert.ExecuteNonQuery();
                    commandInsert.CommandText = string.Format(strSQL5, strMonth);
                    intLow = commandInsert.ExecuteNonQuery();
                    commandInsert.CommandText = string.Format(strSQL6, strMonth);
                    intLow = commandInsert.ExecuteNonQuery();
                    commandInsert.CommandText = string.Format(strSQL7, strMonth);
                    intLow = commandInsert.ExecuteNonQuery();
                    commandInsert.CommandText = string.Format(strSQL8, strMonth);
                    intLow = commandInsert.ExecuteNonQuery();
                    commandInsert.CommandText = string.Format(strSQLInto2, ofd.FileName);
                    commandInsert.ExecuteNonQuery();
                    MessageBox.Show(string.Format(strMessage, strMonth + btnReport.Text, (DateTime.Now - timeStart).TotalSeconds, intLow, intLow / 10), "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (connAccess.State == ConnectionState.Open)
                        connAccess.Close();
                }
            }
        }

        /*
        private void save3DData()
        {
            string strConn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Yaka\Documents\Note.accdb";
            OleDbConnection connAccess = new OleDbConnection(strConn);
            connAccess.Open();
            OleDbCommand commandInsert = new OleDbCommand(strSQLDel, connAccess);

            string strPath = @"G:\3D";
            savefiles(strPath, commandInsert);
            string[] strDirs = Directory.GetDirectories(strPath);
            foreach (string strDir in strDirs)
            {
                savefiles(strDir, commandInsert);
            }
            connAccess.Close();
        }
        private void savefiles(string strPath, OleDbCommand commandInsert)
        {
            string strSQL = "INSERT INTO 3DMovies (FileName,series,ShowYear) VALUES ('{0}','{1}',{2})";
            string[] strfiles = Directory.GetFiles(strPath, @"*.avi");
            string strseries = strPath.Substring(strPath.LastIndexOf("\\") + 1);
            bool is3d = strseries == "3D";
            int intLow;
            foreach (string strfile in strfiles)
            {
                string strname = strfile.Substring(strfile.LastIndexOf("\\") + 1);
                int intindex = strname.IndexOf("20");
                string intYear = "0";
                if (intindex > -1)
                    intYear = strname.Substring(intindex, 4);
                if(is3d)
                {
                    strseries = strname.Substring(0, intindex-1);
                }
                commandInsert.CommandText = string.Format(strSQL, strname, strseries, intYear);
                intLow = commandInsert.ExecuteNonQuery();
            }
        }
        */
        private void btnCalculate_Click(object sender, EventArgs e)
        {
         //   MessageBox.Show("TODO");
            string strWeb = "http://wx.ymiot.net/dwz?tk=pbdgd9qh&tc=158002148{0}";
            //  string invalid = "https://alipay.ymiot.net/SysMsg/?type=2&message=%E4%BC%98%E6%83%A0%E5%88%B8%E6%97%A0%E6%95%88%E6%88%96%E5%B7%B2%E8%BF%87%E6%9C%9F";
            string strValid = "/Coupon/CouponReceive";
            long startNo = 3084845056, endNo = 3698969088; 
            HttpWebRequest LoginReq;
            HttpWebResponse LoginRes;
            string strAddress;
            try
            {
                do
                {
                    startNo += 65536;
                    txtNo.Text = startNo.ToString();
                    txtNo.Refresh();
                    LoginReq = (HttpWebRequest)WebRequest.Create(string.Format(strWeb, startNo));
                    LoginRes = (HttpWebResponse)LoginReq.GetResponse();
                    strAddress = LoginReq.Address.AbsolutePath;
                    LoginRes.Close();
                } while (strAddress != strValid && startNo <= endNo);
                MessageBox.Show(txtNo.Text);
            }
            catch (WebException ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if(DateTime.Now.Minute == 30)
            {
                if (DateTime.Now.Hour == 11)
                {
                    MessageBox.Show(strUserName + strMessageLunch, "整点播报", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else if (DateTime.Now.Hour == 18)
                {
                    MessageBox.Show(strUserName + strMessageOffduty, "整点播报", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else
                    MessageBox.Show(strUserName + strMessageHours, "整点播报", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (DateTime.Now.Minute == 35)
            {
                if (DateTime.Now.Hour == 18)
                {
                    // Process.Start("shutdown.exe", "-s");
                    LockWorkStation();
                }
                else if (DateTime.Now.Hour == 11)
                {
                    LockWorkStation();
                    //Process.Start("shutdown.exe", "-l");
                }
            }

        }

        private void MainForm_Load(object sender, EventArgs e)
        {/*
            RegistryKey R_local = Registry.CurrentUser;
            RegistryKey R_run = R_local.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run");
            R_run.SetValue("Salary4Zhen", Application.ExecutablePath); 
            R_run.Close();
            R_local.Close*/
           // if (Environment.UserName != strUserName)
            strUserName = Environment.UserName;
            //CheckYaka();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = strFilter;
            if(ofd.ShowDialog() == DialogResult.OK)
            {
                OleDbConnection conn = new OleDbConnection(string.Format(strConn, Application.StartupPath));
                
                try
                {
                    DateTime timeStart = DateTime.Now;
                    conn.Open(); 
                    OleDbCommand commandInsert = new OleDbCommand(string.Format(strSQLDel0, strMonth), conn);
                    int intLow = commandInsert.ExecuteNonQuery();
                    commandInsert.CommandText = string.Format(strSQLImport, strMonth, ofd.FileName);
                    intLow = commandInsert.ExecuteNonQuery();
                    MessageBox.Show(string.Format(strMessage, strMonth + btnImport.Text, (DateTime.Now - timeStart).TotalSeconds, intLow, intLow / 100), "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (conn.State == ConnectionState.Open)
                        conn.Close();
                }
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = strFilter;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                OleDbConnection conn = new OleDbConnection(string.Format(strConn, Application.StartupPath));
                try
                {
                    DateTime timeStart = DateTime.Now;
                    conn.Open();
                    OleDbCommand commandInsert = new OleDbCommand(string.Format(strSQLExport, ofd.FileName, strMonth), conn);
                    int intLow = commandInsert.ExecuteNonQuery();
                    MessageBox.Show(string.Format(strMessage, strMonth + btnExport.Text, (DateTime.Now - timeStart).TotalSeconds, intLow, intLow / 100), "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (conn.State == ConnectionState.Open)
                        conn.Close();
                }

            }
        }

        private void listMonth_SelectedValueChanged(object sender, EventArgs e)
        {
            strMonth = (radioHuzhou.Checked ? 'H' + listMonth.Text : listMonth.Text);
        }

      
    }
}
