using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using GoumangToolKit;
using MongoDB.Driver;
using MongoDB.Bson;
using MongoDB.Driver.Builders;
using System.Text.RegularExpressions;

namespace DBQuery
{
    public class Entity
    {
        public ObjectId Id { get; set; }
        public string FileName { get; set; }
        public string InsertDate { get; set; }
        public string FilePath { get; set; }
        public string Rev { get; set; }
        public string Extention { get; set; }

    }
    public static  class QueryParts
    {
        private static string[] QueryItems = new string[]
{
                    "查询图号",
                    "有效地址",
                    "旧版地址",
                    "有效版次",
                    "下级装配号",
                    "名称",
                    "EBOM数量",
                    "级别",
                    "工作包",
                    "工作包名称",
                    "有效架次",
                    "最后发件架次",
                    "配套单机数",
                    "库房结存",
                    "工位号",
                    "库存构型号",
                    "有效构型号",
                    "EBOM构型号",
                    "材料",
                    "分工",
                    "Href"
};
       public static string client = (string)(localMethod.GetConfigValue("MONGO_URI", "PartDBCfg.py"));
       public static string database = (string)(localMethod.GetConfigValue("MONGO_DATABASE", "PartDBCfg.py"));
        public static string col = (string)(localMethod.GetConfigValue("PART_MONGO_COLNAME", "PartDBCfg.py"));
        public static string FTP_col = (string)(localMethod.GetConfigValue("FTP_PART_MONGO_COLNAME", "PartDBCfg.py"));
      static  MongoMethod4 mmlocal = new MongoMethod4(client, database, col);
      static  MongoMethod4 mmFTP = new MongoMethod4(client, database, FTP_col);
        public static  DataTable queryDataList(string extension, List<string> lists)
        {

            if (lists.Count > 0)
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                foreach (var kk in QueryItems)
                {
                    dt.Columns.Add(kk, typeof(string));
                }
                foreach (var pp in lists.AsParallel())
                {


                    var dic = queryData(extension, pp);
                    if (dic != null)
                    {
                        dt.Rows.Add(dic.Values.ToArray());
                    }
                }
                return dt;

            }


            return null;
        }
       public static Dictionary<string, string> queryData(string extension, string partNameQuery)
        {
            

            partNameQuery = partNameQuery.ToUpper();
            // extension is the type of query Date,Drawing,CATPart or pdf
            var rtarry = new Dictionary<string, string>();
            foreach (var kk in QueryItems)
            {
                rtarry.Add(kk, "");
            }
            //输入查询图号
            rtarry["查询图号"] = partNameQuery.Trim();

            //查询EBOM
            Func<string,bool?> ebomfill = delegate (string partnum)
            {
                StringBuilder strSql = new StringBuilder();
                strSql.Append("select NHA,TITLE,QUANTITY,LEVEL,PARTNUMBER,WORKPACKAGE,MATERIALSPECIFICATIONS from partstate ");
                strSql.Append(string.Format("where PARTNUMBER like '{0}%';", (partnum)));
                var ebomDic = DbHelperSQL.getDicOneRow(strSql.ToString());
                if (ebomDic.Count > 0)
                {
                    rtarry["下级装配号"] = ebomDic["NHA"];
                    rtarry["名称"] = ebomDic["TITLE"];
                    rtarry["EBOM数量"] = ebomDic["QUANTITY"];
                    rtarry["级别"] = ebomDic["LEVEL"];
                    rtarry["工作包"] = ebomDic["WORKPACKAGE"];
                    if (ebomDic["TITLE"].Contains("STRINGER"))
                    {
                        rtarry["材料"] = "ALLOY 2196";
                    }
                    else
                    {
                        if (ebomDic["TITLE"].Contains("SKIN") || (ebomDic["TITLE"].Contains("STRAP")))
                        {
                            rtarry["材料"] = "ALLOY 2198";
                        }
                        else if (ebomDic["TITLE"].Contains("FRAME"))
                        {
                            rtarry["材料"] = "ALLOY 7475";
                        }
                        else
                        {
                            rtarry["材料"] = ebomDic["MATERIALSPECIFICATIONS"];
                        }
                    }

                    if (partNameQuery.Contains("C023"))

                    {
                        rtarry["工作包名称"] = "前机身";
                    }
                    else

                    {
                        if (partNameQuery.Contains("C017"))
                        {
                            rtarry["工作包名称"] = "CS300中机身";


                        }
                        else
                        {
                            if (partNameQuery.Contains("C0132") || partNameQuery.Contains("C0133"))
                            {

                                rtarry["工作包名称"] = "CS100中机身";
                            }
                        }
                    }
                    return true;
                }
                else
                {
                    return null;
                }
            };
            //查询MBOM，返回分工信息
            Action<string> mbomfill = delegate (string partnum)
            {
                if (partnum.Count() > 13)
                {
                    partnum = partnum.Substring(0, 13);
                }
                StringBuilder strSql = new StringBuilder();
                strSql.Append("select DIVISION from MBOM ");
                strSql.Append(string.Format("where PARTNUMBER like '{0}%';", partnum));
                var ebomDic = DbHelperSQL.getDicOneRow(strSql.ToString());
                if (ebomDic.Count > 0)
                {
                    rtarry["分工"] = ebomDic["DIVISION"];

                }
            };
            //查询标准件
            Action<string> SPfill = delegate (string partnum)
            {
                StringBuilder strSql = new StringBuilder();
                strSql.Append("select NHA,TITLE,SUM(QUANTITY) AS QTY,LEVEL,PARTNUMBER,WORKPACKAGE,MATERIALSPECIFICATIONS  from(SELECT NHA,TITLE,QUANTITY,LEVEL,PARTNUMBER,WORKPACKAGE,MATERIALSPECIFICATIONS FROM partstate ");
                strSql.Append(string.Format("where PARTNUMBER LIKE '{0}%'ORDER BY QUANTITY DESC) A", (partnum)));
                var ebomDic = DbHelperSQL.getDicOneRow(strSql.ToString());
                if (ebomDic.Count > 0)
                {
                    rtarry["下级装配号"] = ebomDic["NHA"];
                    rtarry["名称"] = ebomDic["TITLE"];
                    rtarry["EBOM数量"] = ebomDic["QTY"];
                    rtarry["级别"] = ebomDic["LEVEL"];
                    rtarry["工作包"] = ebomDic["WORKPACKAGE"];

                    rtarry["材料"] = ebomDic["MATERIALSPECIFICATIONS"];

                }


            };
            //查询有效数据集填充
            Func<string,string> effectfill = delegate (string partnum)
            {
                //有效数据集给出有效图号、有效版次和有效架次
                //仅查询基本号，匹配有效图号
                StringBuilder strSql_effect = new StringBuilder();
                strSql_effect.Append("select 文件名,图号,文件版次,有效性 from effect_data ");

                if (/*extension == "CATDrawing"||*/ (!rtarry["查询图号"].Contains('-')) || ((extension == "CATDrawing" || extension == "CATProduct") && rtarry["有效构型号"].Contains("001")))
                {
                    strSql_effect.Append(string.Format("where 基本号='{0}' and 数据类型='{1}' order by 文件版次 desc;", partnum.Split('-')[0], extension));
                }
                else
                {
                    strSql_effect.Append(string.Format("where 图号='{0}' and 数据类型='{1}' order by 文件版次 desc;", rtarry["查询图号"], extension));
                }
                var effectDic = DbHelperSQL.getDicOneRow(strSql_effect.ToString());

                if (effectDic.Count > 0)
                {


                   // rtarry["查询图号"] = effectDic["文件名"];
                    rtarry["有效版次"] = effectDic["文件版次"];
                    rtarry["有效架次"] = effectDic["有效性"];
                    //搜索path表，给出文件路径
                    return effectDic["文件名"];


                }
                else
                {
                    return null;
                }
            };
            //查询路径数据库
            Func<string,bool?> pathfill = delegate (string partnum)
              {
                  var c1 = new BsonRegularExpression("/SP/");
                  var c2 = new BsonRegularExpression("/^" + insertspace(partnum.Split('-')[0]) + "/");
                  var c22 = new BsonRegularExpression("/^" + insertspace(partnum) + "/");
                  var c3 = new BsonRegularExpression("/" + extension + "/");
                  var query = new QueryDocument("FileName", c22);
                  var sortBy = SortBy.Descending("Rev");

                  var dt_path = mmlocal.collection.Find(query).SetSortOrder(sortBy);


                  //如果Z盘查询不到，则在FTP中查询
                  if (dt_path.Count() == 0)
                  {

                      query = new QueryDocument("FileName", c22);
                      sortBy = SortBy.Descending("Rev");

                      dt_path = mmFTP.collection.Find(query).SetSortOrder(sortBy);

                  }
                  //再次判断
                  if (dt_path.Count() != 0)
                  {
                      //添加路径
                      string strHyperlinks = dt_path.First()["FilePath"].AsString;

                      rtarry["有效地址"] = "=HYPERLINK(\"" + strHyperlinks + "\", \"" + strHyperlinks + "\")"; ;
                      rtarry["Href"] = strHyperlinks;

                  }
                  //无有效地址时，使用旧版地址
                  else
                  {

                      var ba = new BsonArray();
                      var nba = new BsonArray();
                      nba.Add(new BsonDocument { { "FileName", c1 } });

                      ba.Add(new BsonDocument { { "FileName", c2 } });
                      ba.Add(new BsonDocument { { "$nor", nba } });
                      ba.Add(new BsonDocument { { "Extention", c3 } });

                      query = new QueryDocument { { "$and", ba } };

                      dt_path = mmlocal.collection.Find(query).SetSortOrder(sortBy);


                      if (dt_path.Count() != 0)
                      {
                          string strHyperlinks = dt_path.First()["FilePath"].AsString;

                          rtarry["旧版地址"] = "=HYPERLINK(\"" + strHyperlinks + "\", \"" + strHyperlinks + "\")";
                          rtarry["Href"] = strHyperlinks;
                      }
                      else
                      {
                          return false;


                      }
                    

                  }
                  return null;
              };
            //查询库存表
            Action<string> storefill = delegate (string partnum)
            {
                //查询库存表
                StringBuilder strSql_store = new StringBuilder();
                strSql_store.Append(string.Format("select 最后架次,单机数,结存,工位号 from store_state where 零件号 like '{0}%';", rtarry["查询图号"]));
                var storeDic = DbHelperSQL.getDicOneRow(strSql_store.ToString());
                if (storeDic.Count > 0)
                {
                    rtarry["最后发件架次"] = storeDic["最后架次"];
                    rtarry["配套单机数"] = storeDic["单机数"];
                    rtarry["库房结存"] = storeDic["结存"];
                    rtarry["工位号"] = storeDic["工位号"];
                }
            };
            //查询最新构型号
            Func<string, int, Func<string, string>, string> queryBatch = delegate (string partNameTrunk, int basenum, Func<string, string> queryStr)
            {
                var minnum = ((basenum - 6) < 0) ? 0 : (basenum - 6);
                for (int kk = basenum+4; kk >=minnum; kk -=2)
                {

                    string newname = partNameTrunk + "-" + kk.ToString().PadLeft(3, '0');
                    List<string> tempgouxing = DbHelperSQL.getlist(queryStr(newname));
                    if (tempgouxing.Count() != 0)
                    {
                        string finalname = tempgouxing.First().ToString();
                        return finalname.Split('-')[1].Trim();


                    }

                }

                return null;
            };
            //填充构型号
            Func<string,int,string> revfill = delegate (string partNameTrunk, int basenum)
            {
                
                rtarry["有效构型号"] = queryBatch(partNameTrunk, basenum, delegate (string newname)
                {
                    return "select 图号 from effect_data where 图号 like '" + newname + "%';";
                });


                rtarry["EBOM构型号"] = queryBatch(partNameTrunk, basenum, delegate (string newname)
                {
                    return "select PARTNUMBER from partstate where PARTNUMBER like '" + newname + "%';";
                });
                rtarry["库存构型号"] = queryBatch(partNameTrunk, basenum, delegate (string newname)
                {
                    return "select 零件号 from store_state where 零件号 like '" + newname + "%';";
                });


                return rtarry["有效构型号"] ?? rtarry["EBOM构型号"] ?? rtarry["库存构型号"];

            };


            //temp filter the standard parts
            //标准件时
            if (partNameQuery.First() != 'C')
            {
                SPfill(rtarry["查询图号"]);
                pathfill(rtarry["查询图号"]);

                return rtarry;
            }
            else
            {
                var nameFormat = partNameQuery.Split('-');
                string partNameTrunk = nameFormat[0].Trim();
                var tmct = nameFormat.Count();
                if ((partNameQuery.Count() > 15 || partNameQuery.Count() < 9) && tmct != 2)
                {
                    //未知的零件或其他
                    //15为最大字符，如C01323100-N0001
                    ebomfill(rtarry["查询图号"]);
                    mbomfill(rtarry["查询图号"]);
                    pathfill(rtarry["查询图号"]);
                    return rtarry;
                }
                else
                {
                    mbomfill(rtarry["查询图号"]);

                    //带有构型号的情况
                    if (tmct == 2)
                    {


                        string partRev = nameFormat[1].Trim();
                        var f = ebomfill(rtarry["查询图号"]) ?? ebomfill(partNameTrunk);

                        //-N00X的情况
                        if (partRev.Count() != 3)
                        {

                            pathfill(effectfill(rtarry["查询图号"]) ?? rtarry["查询图号"]);

                            return rtarry;

                        }
                        //正常 基本号+构型号情况
                        else
                        {
                            //填充构型号
                            revfill(partNameTrunk, System.Convert.ToInt16(partRev));
                            //填充库存构型号
                            storefill(partNameTrunk + "-" + rtarry["库存构型号"]);
                            //在有效数据集中查找到该号
                            var filename = effectfill(rtarry["查询图号"]);
                            if (filename != null)
                            {
                                pathfill(rtarry["查询图号"]);


                            }

                            else
                            {
                                //有效数据集中未找到,定位新的构型号
                                string newpartnum = partNameTrunk + "-" + rtarry["有效构型号"];
                                var r = (pathfill(effectfill(newpartnum) ?? newpartnum)) ?? pathfill(rtarry["查询图号"]);



                            }

                        }

                    }
                    //不带有构型号的情况
                    else
                    {
                        //填充构型号
                        revfill(partNameTrunk, 6);
                        //如果是零件
                        if (extension == "CATPart")
                        {

                            string newpartnum = partNameTrunk + "-" + rtarry["有效构型号"];
                            var r = (pathfill(effectfill(newpartnum) ?? newpartnum)) ?? pathfill(rtarry["查询图号"]);
                            ebomfill(partNameTrunk + "-" + rtarry["EBOM构型号"]);
                            storefill(partNameTrunk + "-" + rtarry["库存构型号"]);




                        }
                        else
                        //如果不是零件
                        {
                            ebomfill(partNameTrunk + "-" + rtarry["EBOM构型号"]);

                            storefill(partNameTrunk + "-" + rtarry["库存构型号"]);
                            pathfill(effectfill(rtarry["查询图号"]) ?? rtarry["查询图号"]);


                        }




                    }


                }

            }



            return rtarry;
        }



        private static string insertspace(string aa)
        {
            if (!aa.Contains(" "))
            {


                if (aa.Contains("--"))
                {
                    aa = aa.Replace("--", " --");
                }
                else
                {
                    aa = aa.Replace("-", " -");
                }
            }
            return aa;
        }




    }
}
