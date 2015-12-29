//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using GoumangToolKit;


namespace SACISearcher
{
    public partial class MainSearchForm : Form
    {

        IOrderedEnumerable<FileInfo> files = null;
        //  List<string> abc=new List<string>();
        List<string> part = new List<string>();
        List<string> drawing = new List<string>();
        List<string> xlss = new List<string>();
        System.Data.DataTable testExcel = null;
        string path1 = "";
        string path2 = "";
        string path3 = "";
        System.Data.DataTable dt;



        public MainSearchForm()
        {
            InitializeComponent();
           
        }

        private void work(string extension, List<string> lists)
        {
            var dt = DBQuery.QueryParts.queryDataList(extension, lists);
            var hrefarray = from pp in dt.AsEnumerable()
                            let hs = pp["Href"].ToString()
                            where hs != ""
                            select hs;

            checkedListBox1.Items.AddRange(hrefarray.ToArray());
          dynamic wSheet=  OFFICE_Method.excelMethod.SaveDataTableToExcel(dt);
            wSheet.Range["C:D"].ColumnWidth = 15;

        }



        private void button1_Click(object sender, EventArgs e)
        {
            //Apply new connect string
           // string org = DbHelperSQL.connectionString;
          //  DbHelperSQL.connectionString = "Database='state';Data Source='192.168.3.32';User Id='partquery';Password='222222';CharSet = utf8";
            List<string> lists = new List<string>();

            lists = richTextBox1.Text.Split(new Char[2] { '\r', '\n' }, System.StringSplitOptions.RemoveEmptyEntries).ToList();

       
          if(  radioButton2.Checked)
          {
              work("CATPart", lists);
          }
                else
          {
              if(  radioButton1.Checked)
              {
                  work("CATDrawing", lists);

              }
              else
              {
                  if(radioButton3.Checked)
                  {
                      work("CATProduct", lists);

                  }
                  else
                  {
                      work("pdf", lists);

                  }

              }
          }
          //  DbHelperSQL.connectionString = org;

        }


        private void findexcel(string path)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            try
            {

                files = dir.GetFiles().OrderBy(i => i.CreationTime);


                foreach (System.IO.FileInfo fi in files)
                {




                    if (fi.Extension == ".xls" || fi.Extension == ".xlsx")
                    {
                         if (!fi.Name.Contains("~"))
                            {
                                xlss.Add(fi.Name);
                            }
                        
                      

                    }

                }

            }

            catch
            {
                MessageBox.Show("Wrong Path!");
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = "";
        }

        private void button3_Click(object sender, EventArgs e)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select PARTNUMBER from partstate ");
            strSql.Append(string.Format("where TITLE like 'PANEL%' or TITLE like '\"PANEL%';"));
            dt = DbHelperSQL.Query(strSql.ToString()).Tables[0];

            //abcgenerate();
            List<string> lists = new List<string>();
            for(int a=0;a<dt.Rows.Count;a++)
            {
                lists.Add(dt.Rows[a][0].ToString().Substring(0, 13));
            }
            work("CATPart", lists);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select PARTNUMBER from partstate ");
            strSql.Append(string.Format("where TITLE like 'STRINGER%' or TITLE like '\"STRINGER%';"));
            dt = DbHelperSQL.Query(strSql.ToString()).Tables[0];

            //abcgenerate();
            List<string> lists = new List<string>();
            for (int a = 0; a < dt.Rows.Count; a++)
            {
                lists.Add(dt.Rows[a][0].ToString().Substring(0, 13));
            }
          work("CATPart", lists);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select PARTNUMBER from partstate ");
            strSql.Append(string.Format("where TITLE like '\"SKIN%' or TITLE like 'SKIN%' ;"));
            dt = DbHelperSQL.Query(strSql.ToString()).Tables[0];

            //abcgenerate();
            List<string> lists = new List<string>();
            for (int a = 0; a < dt.Rows.Count; a++)
            {
                lists.Add(dt.Rows[a][0].ToString().Substring(0,13));
            }
          work("CATPart", lists);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select PARTNUMBER from partstate ");
            strSql.Append(string.Format("where TITLE like '\"FRAME%' or TITLE like 'FRAME%' ;"));
            dt = DbHelperSQL.Query(strSql.ToString()).Tables[0];

            //abcgenerate();
            List<string> lists = new List<string>();
            for (int a = 0; a < dt.Rows.Count; a++)
            {
                lists.Add(dt.Rows[a][0].ToString().Substring(0, 13));
            }
          work("CATPart", lists);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if(textBox1.Text=="")
            {
                MessageBox.Show("请重新输入！");

            }
            else
            { 
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select PARTNUMBER from partstate ");
            strSql.Append(string.Format("where NHA like '{0}%' and PARTNUMBER like 'C0%';", textBox1.Text));
            dt = DbHelperSQL.Query(strSql.ToString()).Tables[0];

            //abcgenerate();
            List<string> lists = new List<string>();
            for (int a = 0; a < dt.Rows.Count; a++)
            {
                lists.Add(dt.Rows[a][0].ToString().Substring(0, 13));
            }
          work("CATPart", lists);
                }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (select_all.Checked)
            {
                for (int j = 0; j < checkedListBox1.Items.Count; j++)
                    checkedListBox1.SetItemChecked(j, true);
            }
            else
            {
                for (int j = 0; j < checkedListBox1.Items.Count; j++)
                    checkedListBox1.SetItemChecked(j, false);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.SelectedPath = "";// 设置默认路径
            DialogResult ret = fbd.ShowDialog();
            //    string strCollected = string.Empty;


            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {

                if (checkedListBox1.GetItemChecked(i))
                {

                    string filename = checkedListBox1.GetItemText(checkedListBox1.Items[i]);
                    if (!filename.Contains("ftp"))
                    {
                        string[] tempnamestr = filename.Split('\\');
                        string shortname = tempnamestr[tempnamestr.Count() - 1];
                        File.Copy(filename, fbd.SelectedPath + "\\" + shortname, true);
                        // MessageBox.Show("下载完成");
                    }
                    else
                    {
                        string[] tempnamestr = filename.Split('/');
                        string shortname = tempnamestr[tempnamestr.Count() - 1];

                        string FileName = fbd.SelectedPath + "\\" + shortname;
                        string FileNameftp = filename;
                        int allbye = this.GetFtpFileSize(FileNameftp);
                        //创建一个文件流
                        FileStream fs = null;
                        Stream responseStream = null;
                        try
                        {

                            var saciftp = new FtpOperation();



                            //获取一个请求响应对象
                            FtpWebResponse response = saciftp.Download(filename);

                            //获取请求的响应流
                            responseStream = response.GetResponseStream();

                            //判断本地文件是否存在，如果存在，则打开和重写本地文件

                            if (File.Exists(FileName))
                            {

                                fs = File.Open(FileName, FileMode.Open, FileAccess.ReadWrite);

                            }

                            //判断本地文件是否存在，如果不存在，则创建本地文件
                            else
                            {
                                fs = File.Create(FileName);
                            }

                            if (fs != null)
                            {

                                int buffer_count = 65536;
                                byte[] buffer = new byte[buffer_count];
                                int size = 0;
                                int startbye = 0;

                                progressBar1.Maximum = allbye;
                                progressBar1.Minimum = 0;
                                progressBar1.Visible = true;
                                // this.lbl_ftpStakt.Visible = true;
                                while ((size = responseStream.Read(buffer, 0, buffer_count)) > 0)
                                {
                                    fs.Write(buffer, 0, size);
                                    startbye += size;
                                    progressBar1.Value = startbye;

                                    label4.Text = "已下载:" + (int)(startbye / 1024) + "KB/" + "总长度:" + (int)(allbye / 1024) + "KB" + " " + " 文件名:" + FileNameftp;
                                    System.Windows.Forms.Application.DoEvents();
                                }
                                fs.Flush();
                                fs.Close();
                                responseStream.Close();
                            }
                        }
                        finally
                        {
                            if (fs != null)
                                fs.Close();
                            if (responseStream != null)
                                responseStream.Close();
                        }





                    


            }
            

           
                }

              
            }
            MessageBox.Show("下载已完成,保存在：" + fbd.SelectedPath);
           
        }
        private int GetFtpFileSize(string fileNameftp)
        {
            string ftpUserID = "saciftp";
            string ftpPassword = "saciftp_C1107";
            FtpWebRequest reqFTP;
            int fileSize = 0;
            try
            {
                
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(fileNameftp));

                reqFTP.Credentials = new NetworkCredential(ftpUserID, ftpPassword);

                reqFTP.Method = WebRequestMethods.Ftp.GetFileSize;

                reqFTP.UseBinary = true;
                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                Stream ftpStream = response.GetResponseStream();
                fileSize = (int)response.ContentLength;

                ftpStream.Close();
                response.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return fileSize;
        }
        private void zlistftp(string path)
        {
            string ftpURI = path;

            string ftpUserID = "saciftp";
            string ftpPassword = "saciftp_C1107";

            try
            {
                StringBuilder result = new StringBuilder();
                FtpWebRequest ftp;
                ftp = (FtpWebRequest)FtpWebRequest.Create(new Uri(ftpURI));
                ftp.Credentials = new NetworkCredential(ftpUserID, ftpPassword);
                ftp.Method = WebRequestMethods.Ftp.ListDirectory;
                WebResponse response = ftp.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.Default);
                 string line = reader.ReadLine();

                 while (line != null)
                {
                    result.Append(line);
                    if (line.Contains(".CATDrawing"))
                    {

                        drawing.Add(line);

                    }
                    else if (line.Contains(".CATPart"))
                    {
                        part.Add(line);
                    }

                    line = reader.ReadLine();
                }


            }

            

            catch
            {
                MessageBox.Show("Wrong Path!");
            }
        }



        private void button11_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("请重新输入！");

            }
            else
            {
                StringBuilder strSql = new StringBuilder();
                strSql.Append("select PARTNUMBER from partstate ");
                strSql.Append(string.Format("where NHA like '{0}%' and PARTNUMBER not like 'C0%';", textBox1.Text));
                dt = DbHelperSQL.Query(strSql.ToString()).Tables[0];

                //abcgenerate();
                List<string> lists = new List<string>();
                for (int a = 0; a < dt.Rows.Count; a++)
                {
                    lists.Add(dt.Rows[a][0].ToString());
                }
              work("CATPart", lists);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {

        }




        private void 库存查询ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 f = new Form2();
            f.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            checkedListBox1.Items.Clear();
        }

        private void fTP管理ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void MainSearchForm_Load(object sender, EventArgs e)
        {

        }
    }
}
