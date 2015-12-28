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
            Action<string> ebomfill = delegate (string partnum)
            {
                StringBuilder strSql = new StringBuilder();
                strSql.Append("select NHA,TITLE,QUANTITY,LEVEL,PARTNUMBER,WORKPACKAGE,MATERIALSPECIFICATIONS from partstate ");
                strSql.Append(string.Format("where PARTNUMBER='{0}';", (partnum)));
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

            //temp filter the standard parts
            if (partNameQuery.First() != 'C')
            {
                SPfill(rtarry["查询图号"]);
                return rtarry;
            }
            else
            {
                if (partNameQuery.Count() > 13 || partNameQuery.Count() < 9)
                {
                    ebomfill(rtarry["查询图号"]);
                    mbomfill(rtarry["查询图号"]);
                    return rtarry;
                }

            }




            var nameFormat = partNameQuery.Split('-');
            string partNameTrunk = nameFormat[0].Trim();
            string partRev;

            //查询最新构型号
            Func<int, int, Func<string, string>, string> queryBatch = delegate (int maxrev, int step, Func<string, string> queryStr)
            {
                for (int kk = maxrev; kk > 0; kk -= step)
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


            if (nameFormat.Count() > 1)
            {
                partRev = nameFormat[1].Trim();

                if (partRev.Count() != 3)
                {
                    ebomfill(rtarry["查询图号"]);
                    mbomfill(rtarry["查询图号"]);
                    rtarry["下级装配号"] = "";
                    return rtarry;

                }

                int partRevint = Convert.ToInt32(partRev);

                //查询有效数据集,确定有效数据集中最新的构型号


                rtarry["有效构型号"] = queryBatch(partRevint + 6, 2, delegate (string newname)
                {
                    return "select 图号 from effect_data where 图号 like '" + newname + "%';";
                }

                );
                //查询ebom,确定ebom中最新的构型号


                rtarry["EBOM构型号"] = queryBatch(partRevint + 6, 2, delegate (string newname)
                {
                    return "select PARTNUMBER from partstate where PARTNUMBER like '" + newname + "%';";
                });
                //查询库存表,确定库存中最新的构型号

                rtarry["库存构型号"] = queryBatch(partRevint + 6, 2, delegate (string newname)
                {
                    return "select 零件号 from store_state where 零件号 like '" + newname + "%';";
                });






            }
            else
            {
                //如果是零件，则查询有效数据集，已最新的构型号作为构型号
                if (extension == "CATPart")
                {

                    rtarry["有效构型号"] = queryBatch(10, 1, delegate (string newname)
                    {
                        return "select 图号 from effect_data where 图号 like '" + newname + "%';";
                    });


                    rtarry["EBOM构型号"] = queryBatch(10, 1, delegate (string newname)
                    {
                        return "select PARTNUMBER from partstate where PARTNUMBER like '" + newname + "%';";
                    });
                    rtarry["库存构型号"] = queryBatch(10, 1, delegate (string newname)
                    {
                        return "select 零件号 from store_state where 零件号 like '" + newname + "%';";
                    });

                    rtarry["查询图号"] = partNameTrunk + "-" + rtarry["有效构型号"];



                }
                else
                {

                    rtarry["查询图号"] = partNameTrunk;
                    rtarry["EBOM构型号"] = "001";

                }
            }

            //查询ebom,填充信息
            ebomfill(partNameTrunk + "-" + rtarry["EBOM构型号"]);
            mbomfill(partNameTrunk + "-" + rtarry["EBOM构型号"]);
            //查询有效数据集,填充信息

            //有效数据集给出有效图号、有效版次和有效架次
            //仅查询基本号，匹配有效图号
            StringBuilder strSql_effect = new StringBuilder();
            strSql_effect.Append("select 文件名,图号,文件版次,有效性 from effect_data ");

            if (/*extension == "CATDrawing"||*/ !rtarry["查询图号"].Contains('-'))
            {
                strSql_effect.Append(string.Format("where 基本号='{0}' and 数据类型='{1}' order by 文件版次 desc;", partNameTrunk, extension));
            }
            else
            {
                strSql_effect.Append(string.Format("where 图号='{0}' and 数据类型='{1}' order by 文件版次 desc;", rtarry["查询图号"], extension));
            }
            var effectDic = DbHelperSQL.getDicOneRow(strSql_effect.ToString());

            if (effectDic.Count > 0)
            {
                rtarry["有效版次"] = effectDic["文件版次"];

                rtarry["有效架次"] = effectDic["有效性"];
                //搜索path表，给出文件路径
                MongoMethod4 mmlocal = new MongoMethod4(client, database, col);
                 MongoMethod4 mmFTP = new MongoMethod4(client, database, FTP_col);
                var c1 = new BsonRegularExpression("/SP/");
                var c2 = new BsonRegularExpression("/^" + insertspace(partNameTrunk) + "/");
                var c22 = new BsonRegularExpression("/^" + insertspace(effectDic["文件名"]) + "/");
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
                    var nba= new BsonArray();
                    nba.Add(new BsonDocument { { "FileName", c1 } });

                    ba.Add(new BsonDocument { { "FileName", c2 } });
                    ba.Add( new BsonDocument { { "$nor",nba }});
                    ba.Add(new BsonDocument {  { "Extention",c3} });

                    query = new QueryDocument { { "$and", ba } };

                    dt_path = mmlocal.collection.Find(query).SetSortOrder(sortBy);


                    if (dt_path.Count()!= 0)
                    {
                        string strHyperlinks = dt_path.First()["FilePath"].AsString;

                        rtarry["旧版地址"] = "=HYPERLINK(\"" + strHyperlinks + "\", \"" + strHyperlinks + "\")";
                        rtarry["Href"] = strHyperlinks;
                    }
                    else
                    {
                        throw new Exception("请确认是否选择了文件的正确类型！");


                    }

                }
                
            }



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
