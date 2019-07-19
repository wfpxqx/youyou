using QianSheng_Data_Export.Common.CommDB;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace QianSheng_Data_Export.Common.BLL
{
    /// <summary>
    /// 获取要导出的数据
    /// </summary>
    public class GetExportData
    {
        #region  获取凭证   
        /// <summary>
        /// 获取凭证
        /// </summary>
        /// <returns></returns>
        public static DataTable GetVoucher(int iyear)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("KJQJ", Type.GetType("System.Int32"));//会计期间
            dt.Columns.Add("PZH", Type.GetType("System.String"));//凭证号
            dt.Columns.Add("PZZ", Type.GetType("System.String"));
            dt.Columns.Add("FLH", Type.GetType("System.Int32"));//分录号
            dt.Columns.Add("PZ_RQ", Type.GetType("System.String"));
            dt.Columns.Add("FLZY", Type.GetType("System.String"));//分录摘要
            dt.Columns.Add("KM_DM", Type.GetType("System.String"));//科目代码
            dt.Columns.Add("JFJE", Type.GetType("System.String"));//借方金额（人民币）
            dt.Columns.Add("DFJE", Type.GetType("System.String"));//贷方金额（人民币）
            dt.Columns.Add("WBBZ", Type.GetType("System.String"));
            dt.Columns.Add("WBJFJE", Type.GetType("System.String"));//外币借方金额（人民币）
            dt.Columns.Add("WBDFJE", Type.GetType("System.String"));//外币贷方金额（人民币）
            dt.Columns.Add("SHR", Type.GetType("System.String"));//审核人
            dt.Columns.Add("ZDR", Type.GetType("System.String"));//制单人
            dt.Columns.Add("JZR", Type.GetType("System.String"));//记账人
            dt.Columns.Add("CN", Type.GetType("System.String"));//出纳
            dt.Columns.Add("FJS", Type.GetType("System.Int32"));//附件数
            StringBuilder sql = new StringBuilder();
            sql.AppendLine("SELECT iperiod,ino_id,csign,inid,dbill_date,cdigest,ccode,md,mc,cexch_name,md_f,mc_f,ccheck,cbill,cbook,ccashier,idoc FROM GL_accvouch ");
            sql.AppendLine(string.Format("where iyear={0}", iyear));
            sql.AppendLine(" and iflag IS NULL and ino_id IS NOT NULL");
            DataTable dtResults = DbOperation.GetDataTable(sql.ToString());
            DataRow drTo = null;
            foreach (DataRow drFrom in dtResults.Rows)
            {
                drTo = dt.NewRow(); 
                drTo["KJQJ"] = drFrom["iperiod"];
                drTo["PZH"] = drFrom["ino_id"];
                drTo["PZZ"] =  drFrom["csign"];//"记帐凭证"; //凭证字
                drTo["FLH"] = drFrom["inid"];
                drTo["PZ_RQ"] = ((DateTime)drFrom["dbill_date"]).ToString("yyyyMMdd");//凭证日期(8位20130106)对制单日期
                drTo["FLZY"] =dealCdigest( drFrom["cdigest"].ToString ());
                drTo["KM_DM"] = drFrom["ccode"];
                drTo["JFJE"] = drFrom["md"].ToString ().TrimEnd('0').TrimEnd('.');
                drTo["DFJE"] = drFrom["mc"].ToString().TrimEnd('0').TrimEnd('.');
                drTo["WBBZ"] = drFrom["cexch_name"];//外币币种名称
                drTo["WBJFJE"] = drFrom["md_f"].ToString().TrimEnd('0').TrimEnd('.').TrimEnd('0'); 
                drTo["WBDFJE"] = drFrom["mc_f"].ToString().TrimEnd('0').TrimEnd('.').TrimEnd('0');
                drTo["SHR"] = drFrom["ccheck"];
                drTo["ZDR"] = drFrom["cbill"];
                drTo["JZR"] = drFrom["cbook"];
                drTo["CN"] = drFrom["ccashier"];
                drTo["FJS"] = drFrom["idoc"];// (U861)  附单据数 
                dt.Rows.Add(drTo);
            }
            return dt;
        }
        /// <summary>
        /// 凭证摘要字段中不能存在“回车符”、“换行符”和“英文双引号”
        /// </summary>
        /// <param name="cdigest"></param>
        /// <returns></returns>
        private static string dealCdigest(string cdigest)
        {
            string result;
            string ignore = "[\r\n\"]";//需要替换的符号\t
            result = Regex.Replace(cdigest, ignore, "");
            return result;
        }
        #endregion
        /// <summary>
        /// 获取帐套信息
        /// </summary>
        /// <returns></returns>
        public static DataTable GetAccount(int iyear)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ZT_DM", Type.GetType("System.String"));// 账套代码
            dt.Columns.Add("ZTMC", Type.GetType("System.String"));// 账套名称", Type.GetType("System.String"));//
            dt.Columns.Add("NSRSBH", Type.GetType("System.String"));// 统一社会信用代码,纳税人识别号
            dt.Columns.Add("ND_DM", Type.GetType("System.String"));// 年度代码(4位2013、2014)
            dt.Columns.Add("JT_DM", Type.GetType("System.String"));// 集团代码
            StringBuilder sql = new StringBuilder();
            sql.AppendLine("SELECT cAcc_Id,cAcc_Name,cUnitTaxNo,iYear,cOrgCode FROM [UFSystem].[dbo].[UA_Account] ");
            sql.AppendLine(string.Format("where iyear={0}", iyear));
            DataTable dtResults = DbOperation.GetDataTable(sql.ToString());
            DataRow drTo = null;
            foreach (DataRow drFrom in dtResults.Rows)
            {
                drTo = dt.NewRow();
                drTo["ZT_DM"] = drFrom["cAcc_Id"];
                drTo["ZTMC"] = drFrom["cAcc_Name"];
                drTo["NSRSBH"] = drFrom["cUnitTaxNo"]; //(U861)  税号 
                drTo["ND_DM"] = drFrom["iYear"];
                drTo["JT_DM"] = drFrom["cOrgCode"];
                dt.Rows.Add(drTo);
            }

            return dt;
        }
        #region 获取会计科目
        /// <summary>
        /// 获取会计科目
        /// </summary>
        /// <returns></returns>
        public static DataTable GetCode(int iyear)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ND_DM", Type.GetType("System.String"));// 年度代码(4位2013、2014)
            dt.Columns.Add("KM_DM", Type.GetType("System.String"));// 科目代码
            dt.Columns.Add("KMMC", Type.GetType("System.String"));// 科目名称
            dt.Columns.Add("KMLX", Type.GetType("System.String"));// 科目类型(1资产2负债3共同4权益5成本6损益7其他9未知)
            dt.Columns.Add("KMFX", Type.GetType("System.String"));// 科目方向(1 - 借方 2 - 贷方 3 - 其他 9 - 未知)
            dt.Columns.Add("FJKM_DM", Type.GetType("System.String"));// 父级科目代码，一级科目的为空
            StringBuilder sql = new StringBuilder();
            sql.AppendLine("SELECT iyear,ccode,ccode_name,cclass,bproperty FROM code ");
            sql.AppendLine(string.Format("where iyear={0}", iyear));
            sql.AppendLine(" order by ccode");

            DataTable dtResults = DbOperation.GetDataTable(sql.ToString());
            DataRow drTo = null;
            string ccode;
            foreach (DataRow drFrom in dtResults.Rows)
            {
                drTo = dt.NewRow();
                ccode = drFrom["ccode"].ToString();
                drTo["ND_DM"] = drFrom["iyear"];
                drTo["KM_DM"] = ccode;
                drTo["KMMC"] = drFrom["ccode_name"];
                drTo["KMLX"] = GetKMLX(drFrom["cclass"].ToString());
                drTo["KMFX"] = (bool)drFrom["bproperty"] ? 1 : 2;//bproperty (U861)  科目性质  1-借方 2-贷方
                drTo["FJKM_DM"] = GetFJKM_DM(ccode);
                dt.Rows.Add(drTo);
            }
            return dt;
        }
        /// <summary>
        /// 获取父级科目代码
        /// </summary>
        /// <param name="ccode"></param>
        /// <returns></returns>
        private static string GetFJKM_DM(string ccode)
        {
            string fjkm="";
            if (ccode.Length>4)
            {
                fjkm = ccode.Substring(0, ccode.Length - 2);
            }
            return fjkm;
        }
        /// <summary>
        /// 获取科目类型名称对应的ID
        /// </summary>
        /// <param name="lxName"></param>
        /// <returns></returns>
        private static  string GetKMLX(string lxName)
        {
            int lxID;
            switch (lxName)
            {
                case "资产":
                    lxID = 1;
                    break;
                case "负债":
                    lxID = 2;
                    break;
                case "共同":
                    lxID = 3;
                    break;
                case "权益":
                    lxID = 4;
                    break;
                case "成本":
                    lxID = 5;
                    break;
                case "损益":
                    lxID = 6;
                    break;
                case "其他":
                    lxID = 7;
                    break;
                default:// "未知":
                    lxID = 9;
                    break;
            }
            return lxID.ToString ();

        }
        #endregion
        /// <summary>
        /// 获取期初余额
        /// </summary>
        /// <returns></returns>
        public static DataTable GetVoucherSum(int iyear)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ND_DM", Type.GetType("System.String"));//年度代码(4位2013、2014)
            dt.Columns.Add("KJQJ", Type.GetType("System.String"));//会计期间（1 - 12）
            dt.Columns.Add("KM_DM", Type.GetType("System.String"));//科目代码
            dt.Columns.Add("QCJFYE", Type.GetType("System.Decimal"));//期初借方余额（人民币）
            dt.Columns.Add("QCDFYE", Type.GetType("System.Decimal"));//期初贷方余额（人民币）
            StringBuilder sql = new StringBuilder();
            sql.AppendLine("SELECT iyear,ccode,iperiod,mb,cbegind_c FROM GL_accsum ");
            sql.AppendLine(string.Format("where iyear={0}", iyear));
            sql.AppendLine(" order by ccode");
            DataTable dtResults = DbOperation.GetDataTable(sql.ToString());
            DataRow drTo = null;
            foreach (DataRow drFrom in dtResults.Rows)
            {
                drTo = dt.NewRow();
                drTo["ND_DM"] = drFrom["iyear"];
                drTo["KJQJ"] = drFrom["iperiod"];
                drTo["KM_DM"] = drFrom["ccode"];
                if (drFrom["cbegind_c"].ToString() == "借")
                {
                    drTo["QCJFYE"] = drFrom["mb"].ToString().TrimEnd('0').TrimEnd('.');  //mb(U861)  金额期初 cbegind_c (U861)  金额期初方向  
                    drTo["QCDFYE"] = 0;//me (U861)  金额期末  cendd_c (U861)  金额期末方向 
                }
                else if (drFrom["cbegind_c"].ToString() == "贷")//还有一种平
                {
                    drTo["QCJFYE"] = 0;
                    drTo["QCDFYE"] = drFrom["mb"].ToString().TrimEnd('0').TrimEnd('.'); 
                }
                dt.Rows.Add(drTo);
            }
            return dt;
        }
    }
}
