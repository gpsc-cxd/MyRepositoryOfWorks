
using Microsoft.International.Converters.PinYinConverter;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OSGeo.OGR;
/**
* 命名空间: 承包地13年与17年shp对比 
* 类 名： Class1
*
* Ver      编写人        变更内容            变更日期
* ──────────────────────────────────────────────────────────────
* V1.0     陈晓东         初版                2018/8/16 16:40:21 
*
* Copyright (c) 2018 SuperSIT. All rights reserved. 
*┌──────────────────────────────────────────────────────────────┐
*│　此技术信息为本公司机密信息，未经本公司书面同意禁止向第三方披露．     |
*│　版权所有：四川旭普信息产业发展有限公司　　　　　　　　　　　　　　   |
*└──────────────────────────────────────────────────────────────┘
*/
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using 承包地13年与17年shp对比;

namespace 承包地13年与17年shp对比
{
    public class Funcs
    {
        /// <summary>
        /// 判断feNew的重心是否在feOld里面
        /// </summary>
        /// <param name="feNew"></param>
        /// <param name="feOld"></param>
        /// <returns></returns>
        public static bool isInside(Feature feNew, Feature feOld)
        {
            Geometry geoNew = feNew.GetGeometryRef().Centroid();
            double aa = geoNew.Area();
            double bb = feOld.GetGeometryRef().Area();
            bool valid = geoNew.IsValid();
            bool isContain = feOld.GetGeometryRef().Contains(geoNew);
            return feOld.GetGeometryRef().Contains(geoNew);
        }
        public static bool isInside(Feature feNew, List<Form1.Features> feOld)
        {
            bool res = false;
            foreach(var fe in feOld)
            {
                if (isInside(feNew, fe.feature))
                {
                    res = true;
                    break;
                }
            }
            return res;
        }
        public static FileStream[] fsPass;
        public static DataTable[] dtPass;
        public static List<Form1.Bsm> bsmPass;
        public static string codes;
        public static void ThreadReadExcel(string[] cbfFiles, out FileStream[] fileStream, out IWorkbook[] workbook, out ISheet[] worksheet, out DataTable[] dataTable)
        {
            #region 读取承包方调查表
            int count = cbfFiles.Length;
            IWorkbook[] wb = new IWorkbook[count];
            ISheet[] ws = new ISheet[count];
            DataTable[] dt = new DataTable[count];
            System.IO.FileStream[] fs = new System.IO.FileStream[count];
            //Thread[] thread = new Thread[count];
            for (int i = 0; i < count; i++)
            {
                ReadExcel(cbfFiles[i], out wb[i], out dt[i], out fs[i]);
            }
            //for (int i = 0; i < thread.Length; i++)
            //{
            //    while (thread[i].IsAlive)
            //        Thread.Sleep(10);
            //}
            for (int i = 0; i < wb.Length; i++)
            {
                ws[i] = wb[i].GetSheetAt(0);//获取第一张表
            }
            fileStream = fs;
            workbook = wb;
            worksheet = ws;
            dataTable = dt;
            #endregion
        }
        private static void ReadExcel(string filePath, out IWorkbook workbook, out DataTable dt, out System.IO.FileStream fs)
        {
            using (fs = new System.IO.FileStream(filePath, FileMode.Open))
            {
                workbook = new HSSFWorkbook(fs);
                dt = new DataTable();
                //读入第一张表到DataTable
                string sheetName = workbook.GetSheetAt(0).SheetName;//获取第一张表名
                dt = getExcel(filePath, sheetName);
                //处理DataTable首尾数据
                TrimDataTable(dt);
                //处理表头
                EditFieldsName(dt, "承包方编码", 1);
            }
        }
        /// <summary>
        /// 读取Excel返回DataTable
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="sheetName">表名</param>
        /// <returns></returns>
        public static DataTable getExcel(string filePath, string sheetName)
        {

            bool hasTitle = false;
            string fileType = Path.GetExtension(filePath);
            using (DataSet ds = new DataSet())
            {
                string strCon = string.Format("Provider=Microsoft.{0}.OLEDB.{1}.0;" +
                                    "Extended Properties=\"Excel {2}.0;HDR=YES;IMEX=1;\";" +
                                    "data source={4};",
                                    (fileType == ".xls" ? "Jet" : "ACE"), (fileType == ".xls" ? 4 : 12), (fileType == ".xls" ? 8 : 12), (hasTitle ? "Yes" : "NO"), filePath);
                string strCom = "SELECT * FROM [" + sheetName + "$]";
                using (OleDbConnection myConn = new OleDbConnection(strCon))
                using (OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn))
                {
                    myConn.Open();
                    myCommand.Fill(ds);
                }
                if (ds == null || ds.Tables.Count <= 0) return null;
                return ds.Tables[0];
            }
        }
        /// <summary>
        /// 截去Datatable数据首尾的空行和合计行
        /// </summary>
        /// <param name="dt"></param>
        private static void TrimDataTable(DataTable dt)
        {
            //先处理首部空行
            while (dt.Rows[0][0].ToString().Trim() == "")
            {
                dt.Rows.RemoveAt(0);
            }
            //处理尾部空行
            int i = dt.Rows.Count - 1;
            while (dt.Rows[i][0].ToString().Trim() == "" || dt.Rows[i][0].ToString().Trim() == "合计")
            {
                dt.Rows.RemoveAt(i);
                i--;
            }
        }

        /// <summary>
        /// 处理DataTable字段名
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="start">查找的第一行第一个字段名称</param>
        /// <param name="headCount">表头字段的行数</param>
        private static void EditFieldsName(DataTable dt, string start, int headCount)
        {
            int st = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][0].ToString().Trim().Replace("\n", "") == start)
                {
                    st = i;
                    break;
                }
            }
            for (int i = st - 1; i >= 0; i--)
            {
                dt.Rows.RemoveAt(i);
            }
            for (int i = 0; i <= headCount - 1; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    int idx = 0;
                    if (dt.Rows[i][j].ToString().Trim() != "")
                    {
                        string name, fname;
                        name = fname = GetFirstSpell(TrimFieldName(dt.Rows[i][j].ToString().Trim()));
                        while (dt.Columns.Contains(name))
                        {
                            idx++;
                            name = fname + idx.ToString();
                        }
                        dt.Columns[j].ColumnName = name;

                    }
                }
            }
            for (int i = headCount - 1; i >= 0; i--)
            {
                dt.Rows.RemoveAt(i);
            }
        }
        /// <summary>
        /// 截去字符串中的各种符号
        /// </summary>
        /// <param name="FieldName"></param>
        /// <returns></returns>
        private static string TrimFieldName(string FieldName)
        {
            Regex reg = new Regex(@"([^\u4e00-\u9fa5a-zA-z0-9\s].*?)"); //去掉一般符号
            string s = reg.Replace(FieldName, "");
            s = s.Replace("\n", "");    //去掉回车符
            //特殊处理“使用”和“所有”造成首字母一样的情况
            //使用->SHY，所有->SY
            s = s.Replace("使用", "使和用");
            //处理四至东南西北字段名太短的问题
            switch (s)
            {
                case "东": s = "宗地四至东"; break;
                case "南": s = "宗地四至南"; break;
                case "西": s = "宗地四至西"; break;
                case "北": s = "宗地四至北"; break;
            }
            return s;
        }
        /// <summary>
        /// 汉字转首字母
        /// </summary>
        /// <param name="strChinese"></param>
        /// <returns></returns>
        public static string GetFirstSpell(string strChinese)
        {
            //NPinyin.Pinyin.GetInitials(strChinese)  有Bug  洺无法识别
            //return NPinyin.Pinyin.GetInitials(strChinese);

            try
            {
                if (strChinese.Length != 0)
                {
                    StringBuilder fullSpell = new StringBuilder();
                    for (int i = 0; i < strChinese.Length; i++)
                    {
                        var chr = strChinese[i];
                        fullSpell.Append(GetSpell(chr)[0]);
                    }

                    return fullSpell.ToString().ToUpper();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("首字母转化出错！" + e.Message);
            }

            return string.Empty;
        }
        private static string GetSpell(char chr)
        {
            var coverchr = NPinyin.Pinyin.GetPinyin(chr);

            bool isChineses = ChineseChar.IsValidChar(coverchr[0]);
            if (isChineses)
            {
                ChineseChar chineseChar = new ChineseChar(coverchr[0]);
                foreach (string value in chineseChar.Pinyins)
                {
                    if (!string.IsNullOrEmpty(value))
                    {
                        return value.Remove(value.Length - 1, 1);
                    }
                }
            }

            return coverchr;

        }
        public static OpenFileDialog myOpenDLG(string filter = "*.*", string title = "",
            bool isRestoreDirectory = false, bool isMultiSelect = true)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Title = title;
            open.Filter = filter;
            open.RestoreDirectory = isRestoreDirectory;
            open.Multiselect = isMultiSelect;
            return open;
        }
        /// <summary>
        /// 将字符串中的数字转成汉字
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string convertNumToCha(string input,bool fan)
        {
            string split = Regex.Replace(input, @"[^0-9]+","");
            string cha = Chinese(split, fan);
            string res = input.Replace(split, cha);
            return res;
        }
        public static string Chinese(object n, bool fan = false)
        {
            string strn = n.ToString();
            string str = "";
            string nn = "零壹贰叁肆伍陆柒捌玖";
            string ln = "零一二三四五六七八九";

            string mm = "  拾佰仟萬拾佰仟亿拾佰仟萬兆拾佰仟萬亿";
            string lm = "  十百千万十百千亿十百千万兆十百千万亿";

            int i = 0;
            while (i < strn.Length)//>>>>>>>>>>>>>>>>出现空格
            {

                int m = int.Parse(strn.Substring(i, 1));
                if (fan)//返回繁体字
                {
                    str += nn.Substring(m, 1);
                    if (lm.Substring(strn.Length - i, 1) != " ")
                    { str += mm.Substring(strn.Length - i, 1); }
                }
                else//返回简体字
                {
                    str += ln.Substring(m, 1);
                    if (lm.Substring(strn.Length - i, 1) != " ")
                    {
                        str += lm.Substring(strn.Length - i, 1);
                    }
                }
                i++;
            }
            if (str.Substring(str.Length - 1) == "零")
            { str = str.Substring(0, str.Length - 1); }
            if (str.Length > 1 && str.Substring(0, 2) == "一十")
            { str = str.Substring(1); }
            if (str.Length > 1 && str.Substring(0, 2) == "壹拾")
            { str = str.Substring(1); }

            return str;
        }
        public static string[] splitCodes(string codes)
        {
            Regex reg = new Regex(@"[0-9]+");
            MatchCollection mc= reg.Matches(codes);
            string[] res = new string[mc.Count];
            for(int i = 0; i < res.Length; i++)
            {
                res[i] = mc[i].Value;
            }
            return res;
        }
    }


}
