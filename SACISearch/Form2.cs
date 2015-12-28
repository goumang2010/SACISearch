
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using GoumangToolKit;
using OFFICE_Method;

namespace SACISearcher
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private string convetUTF8(string str)
        {
           

            byte[] bytes = Encoding.UTF8.GetBytes(str);
            string newstr = Encoding.UTF8.GetString(bytes);
            return newstr;
        }
        private void button1_Click(object sender, EventArgs e)
        {

            
        }

        private void button2_Click(object sender, EventArgs e)
        {

            List<System.Data.DataTable> excelTable=new List<System.Data.DataTable>();
            //List<string> abcname=new List<string>();
            List<string> zhanweilist = new List<string>();

            //string strSql="select aa.站位,aa.工装号,bb.工位号,bb.零件号,bb.名称,bb.最后架次,bb.结存-bb.单机数量 from station aa inner join store_state bb on bb.工位号 like %aa.工位%"
            string strSql = "select 工位 from station where 1";
            //dmm= DbHelperSQL.Query(strSql.ToString()).Tables[0];
            zhanweilist = DbHelperSQL.getlist(strSql.ToString());

            foreach(string p in zhanweilist)
            {

               string biban=  DbHelperSQL.getlist("select 工装名称 from station where 工位='"+p+"'").First();
               if (biban.Contains("壁板"))
               {


                   string strSql2 = "select bb.零件号 as Part,bb.名称 as name,bb.最后架次 as lastFuseNo,bb.单机数 as per,bb.结存 as lastQty  from store_state bb where bb.工位号 like '%" + p + "%'";
                   //List<string> abcname;
                   excelTable.Add(DbHelperSQL.Query(strSql2.ToString()).Tables[0]);

               }
            }
            Dictionary<string, DataTable> excelDic = new Dictionary<string, DataTable>();
            for(int i=0;i<zhanweilist.Count();i++)
            {
                excelDic.Add(zhanweilist[i], excelTable[i]);
            }
        
            excelMethod.SaveDataTableToExcel(excelDic);


           // excelMethod.SaveDataTableToExcel(dmm);
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            string strSql = "select 工装名称 from station where 1";
            listBox1.DataSource = DbHelperSQL.getlist(strSql.ToString());


          //  List<string> jiacilist = new List<string>();
           // List<string> newjiacilist = new List<string>();
           listBox2.DataSource = DbHelperSQL.getlist("select 架次 from store_state_all group by 架次");

        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            System.Data.DataTable dmm;
            string nameSelect="";
            string zhanwei;
          //  listBox1.SelectedItem()
            if (listBox1.SelectedIndex != -1)
            {
                nameSelect = listBox1.SelectedItem.ToString();
            }
            string strSql = "select 工位 from station where 工装名称='" + nameSelect + "'";
            //dmm= DbHelperSQL.Query(strSql.ToString()).Tables[0];
            zhanwei = DbHelperSQL.getlist(strSql.ToString()).First();
            string strSql2 = "select bb.零件号 as Part,bb.名称 as name,bb.最后架次 as lastFuseNo,bb.单机数 as per,bb.结存 as lastQty  from store_state bb where bb.工位号 like '%" + zhanwei + "%'";
            dmm = DbHelperSQL.Query(strSql2.ToString()).Tables[0];
            excelMethod.SaveDataTableToExcel(dmm);
        }

        private void button3_Click(object sender, EventArgs e)
        {
           // string strSql2 = "select a.NHA,b.* from (select bb.零件号 as Part,bb.名称 as name,bb.工位号 as station,bb.最后架次 as lastFuseNo,bb.单机数 as per,bb.结存 as lastQty  from store_state bb where bb.名称='蒙皮') b left join partstate a on a.PARTNUMBER=b.Part";


           // string strSql2 = "select bb.零件号 as Part from store_state bb where bb.名称='蒙皮'and bb.最后架次<>'No'";
            List<string> skinlist=new List<string>();
            List<string> newskinlist = new List<string>();
         
           
              //  string newkk = kk.Substring(0, 9);
            string sql1 = "(select a.工位号,a.最后架次,a.零件号,b.站位 from store_state a left join station b on a.工位号=b.工位  where a.名称='蒙皮' Or a.名称 like 'SKIN%')";
                string sql2 = "(select 工位号,count(*) as 缺件数 from store_state where 单机数>结存 group by 工位号)";
                string sql3 = "select aa.零件号 as skinPart,aa.站位 as LS,aa.工位号 as station,aa.最后架次 as lastFuseNo,bb.缺件数 as LackItemQty from" + sql1 + " aa left join " + sql2 + " bb on aa.工位号=bb.工位号";
                System.Data.DataTable table1;
                table1 = DbHelperSQL.Query(sql3.ToString()).Tables[0];
                dynamic wSheet = excelMethod.SaveDataTableToExcel(table1);
                int editqty = table1.Rows.Count;
                int offsetqty = table1.Columns.Count;
                for (int i = 0; i < editqty; i++)
                {
                    string newkk = table1.Rows[i][0].ToString().Substring(0, 9);
                    string newkkrear=table1.Rows[i][0].ToString().Substring(10, 3);
                    int newkkrearint = Convert.ToInt32(newkkrear);
                    string productpart = "";
                    for (int m = newkkrearint; m > 0; m -= 2)
                    {
                        string rearconvert = m.ToString().PadLeft(3, '0');
                        StringBuilder strSql2 = new StringBuilder();
                        strSql2.Append("select NHA from partstate ");
                        strSql2.Append(string.Format("where PARTNUMBER like '{0}%';", newkk + "-" + rearconvert));
                        System.Data.DataTable dt = DbHelperSQL.Query(strSql2.ToString()).Tables[0];
                        if(dt.Rows.Count!=0)
                        {
                            wSheet.Cells[i + 2, 1 + offsetqty] = dt.Rows[0][0];
                            productpart = dt.Rows[0][0].ToString();
                        }

                    }
                   

                    
                    StringBuilder strSql3 = new StringBuilder();
                    strSql3.Append("select TITLE from partstate ");
                    strSql3.Append(string.Format("where PARTNUMBER like '{0}%';", productpart));
               //DbHelperSQL.getlist(strSql3).First().ToString()
                        //excelTable.Columns[i].ColumnName;
                    wSheet.Cells[i + 2, 2 + offsetqty] = DbHelperSQL.getlist(strSql3.ToString()).First().ToString();
                }
                wSheet.Cells[1, 1 + offsetqty] = "ProductNo";
                wSheet.Cells[1, 2 + offsetqty] = "ProductName";

             //   string sql4 = "( select NHA,TITLE from from partstate where PARTNUMBER like '" + newkk + "%'";
                //excelMethod.SaveDataTableToExcel_offset

                wSheet.Columns.AutoFit();

            //excelMethod.SaveDataTableToExcel(DbHelperSQL.Query(strSql2.ToString()).Tables[0]);
        }

        private void button4_Click(object sender, EventArgs e)
        {
        string strSql2 = "select bb.零件号 as Part,bb.名称 as name,bb.最后架次 as lastFuseNo,bb.单机数 as per,bb.结存 as lastQty  from store_state bb where bb.工位号 like '%" + textBox1.Text + "%'";
           
            
            excelMethod.SaveDataTableToExcel(DbHelperSQL.Query(strSql2.ToString()).Tables[0]);
        
        }

        private void button5_Click(object sender, EventArgs e)
        {
         
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            int jiecuncol;
            if (textBox2.Text != "")
            {
                jiecuncol = System.Convert.ToInt32(textBox2.Text);
            }
            else
            {
                jiecuncol = 91;

            }
            // DbHelperSQL.ExecuteSql("delete from store_state");


            System.Data.DataTable testExcel = null;

            List<string> creatName = new List<string>();

            int k = 0;//一共添加多少条记录




            OpenFileDialog fileDialog = new OpenFileDialog();

            fileDialog.InitialDirectory = "D://";

            fileDialog.Filter = "xls files (*.xls)|*.xls|All files (*.*)|*.*";

            fileDialog.FilterIndex = 1;

            fileDialog.RestoreDirectory = true;

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {

                testExcel = excelMethod.LoadDataFromExcel(fileDialog.FileName);

            }



            //筛选中excel的绩效考核全部信息
            DataRow[] shuju = testExcel.Select();
            //DataRow[] biaotou = testExcel.Select(@"零件号=''");



            foreach (DataRow p in shuju)
            {
                int colcount;
                colcount = jiecuncol - 2;
                string yifajiaci = "No";
                // double[] a = new double[15];

                int danjishu;

                if (p[4].ToString() == "")
                {
                    danjishu = 0;
                }
                else
                {
                    danjishu = System.Convert.ToInt32(p[4].ToString());
                }




                for (int i = colcount; i > 8; i--)
                {
                    string datestr = p[i].ToString();


                    if (datestr.Contains("41") || datestr.Contains("2014"))
                    {
                        int datenum;
                        datenum = -1;
                        try
                        {
                             datenum = System.Convert.ToInt32(datestr);
                        }
                        catch
                        {
                            datenum = -1;
                        }
                        if (datenum > 41000 || datestr.Contains("2014"))
                        { 

                        yifajiaci = testExcel.Columns[i].ColumnName;
                        if (!yifajiaci.Contains("SACI"))
                        {
                            for (int ii = i - 3; ii < i + 2; ii++)
                            {
                                if (testExcel.Columns[ii].ColumnName.Contains("SACI"))
                                {
                                    yifajiaci = testExcel.Columns[ii].ColumnName;

                                    StringBuilder strSqlname = new StringBuilder();
                                    strSqlname.Append("INSERT INTO store_state_all (");

                                    strSqlname.Append("零件号,名称,单机数,工位号,架次,日期");

                                    strSqlname.Append(string.Format(") VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", p[2].ToString(), p[3].ToString(), danjishu, p[5].ToString(), yifajiaci, datestr));



                                    creatName.Add(strSqlname.ToString());
                                    break;
                                    
                                }

                            }

                        }

                        else
                        {
                           StringBuilder strSqlname = new StringBuilder();
                            strSqlname.Append("INSERT INTO store_state_all (");

                            strSqlname.Append("零件号,名称,单机数,工位号,架次,日期");

                            strSqlname.Append(string.Format(") VALUES ('{0}','{1}',{2},'{3}','{4}','{5}')", p[2].ToString(), p[3].ToString(), danjishu, p[5].ToString(), yifajiaci, datestr));



                            creatName.Add(strSqlname.ToString());
                        }
                       // break;
                    }

                }// if 41

               

                //System.Text.Encoding.UTF8.GetString()

                

            }//for 



            //  MessageBox.Show("OK");


            }//for each

            k = DbHelperSQL.ExecuteSqlTran(creatName);
            MessageBox.Show(string.Format("执行成功,增加 '{0}'条记录", k));
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
           int jj = DbHelperSQL.ExecuteSql("delete from store_state_all");

         //   k = DbHelperSQL.ExecuteSqlTran(creatName);
           MessageBox.Show(string.Format("执行成功,清除 '{0}'条记录", jj));
        }

        private void button8_Click(object sender, EventArgs e)
        {
            // string strSql2 = "select bb.零件号 as Part from store_state bb where bb.名称='蒙皮'and bb.最后架次<>'No'";
           // List<string> skinlist = new List<string>();
          //  List<string> jiacilist = new List<string>();
           // List<string> newjiacilist = new List<string>();

            System.Windows.Forms.ListBox.SelectedObjectCollection selectedDoc = this.listBox2.SelectedItems;


           // jiacilist = DbHelperSQL.getlist("select 架次 from store_state_all group by 架次");
         /*   foreach(string pp in jiacilist)
            {
               string npp =pp.Substring(0, 14);
                if(!newjiacilist.Contains(npp))
                {
                    newjiacilist.Add(npp);
                }

            }
            */

            foreach (string nnp in selectedDoc)
            {

            
            //  string newkk = kk.Substring(0, 9);
            string sql1 = "(select a.工位号,a.零件号,b.站位 from store_state a left join station b on a.工位号=b.工位  where a.名称='蒙皮' Or a.名称 like 'SKIN%')";
            string sql2 = "(select 零件号,日期 from store_state_all where 架次 = '"+nnp+"')";
            string sql3 = "select aa.零件号 as skinPart,aa.站位 as LS,aa.工位号 as station,bb.日期 as fajian from" + sql1 + " aa left join " + sql2 + " bb on aa.零件号=bb.零件号";
            System.Data.DataTable table1;
            table1 = DbHelperSQL.Query(sql3.ToString()).Tables[0];
            dynamic wSheet = excelMethod.SaveDataTableToExcel(table1);
            /*  int editqty = table1.Rows.Count;
             int offsetqty = table1.Columns.Count;
             for (int i = 0; i < editqty; i++)
              {
                  string newkk = table1.Rows[i][0].ToString().Substring(0, 9);
                  string newkkrear = table1.Rows[i][0].ToString().Substring(10, 3);
                  int newkkrearint = Convert.ToInt32(newkkrear);
                  string productpart = "";
                  for (int m = newkkrearint; m > 0; m -= 2)
                  {
                      string rearconvert = m.ToString().PadLeft(3, '0');
                      StringBuilder strSql2 = new StringBuilder();
                      strSql2.Append("select NHA from partstate ");
                      strSql2.Append(string.Format("where PARTNUMBER like '{0}%';", newkk + "-" + rearconvert));
                      System.Data.DataTable dt = DbHelperSQL.Query(strSql2.ToString()).Tables[0];
                      if (dt.Rows.Count != 0)
                      {
                          wSheet.Cells[i + 2, 1 + offsetqty] = dt.Rows[0][0];
                          productpart = dt.Rows[0][0].ToString();
                      }

                  }
           



                  StringBuilder strSql3 = new StringBuilder();
                  strSql3.Append("select TITLE from partstate ");
                  strSql3.Append(string.Format("where PARTNUMBER like '{0}%';", productpart));
                  //DbHelperSQL.getlist(strSql3).First().ToString()
                  //excelTable.Columns[i].ColumnName;
                  wSheet.Cells[i + 2, 2 + offsetqty] = DbHelperSQL.getlist(strSql3.ToString()).First().ToString();
              }
              wSheet.Cells[1, 1 + offsetqty] = "ProductNo";
              wSheet.Cells[1, 2 + offsetqty] = "ProductName";
                   * */
            //   string sql4 = "( select NHA,TITLE from from partstate where PARTNUMBER like '" + newkk + "%'";
            //excelMethod.SaveDataTableToExcel_offset

            wSheet.Columns.AutoFit();
              dynamic dateColumn= wSheet.Columns[4];
              dateColumn.NumberFormat = "yyyy-mm-dd";
                //wSheet.
        }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < listBox2.Items.Count; i++)
            {
                listBox2.SelectedIndex = i;
            }

            listBox2.Refresh();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            
                listBox2.ClearSelected();
          

            listBox2.Refresh();
        }
    }
}
