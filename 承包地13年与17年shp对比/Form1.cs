using Microsoft.International.Converters.PinYinConverter;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Linq;
using OSGeo.OGR;
using OSGeo.GDAL;
using System.Drawing;

namespace 承包地13年与17年shp对比
{
    public partial class Form1 : Form
    {
        #region import DLL to check if file is opened...
        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);
        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);
        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        public readonly IntPtr HFILE_ERROR = new IntPtr(-1);
        #endregion
        [DllImport("gdal203.dll", EntryPoint = "OGR_F_GetFieldAsString", CallingConvention = CallingConvention.Cdecl)]
        public extern static IntPtr OGR_F_GetFieldAsString(HandleRef handle, int index);
        public Form1()
        {

            InitializeComponent();
        }
        /// <summary>
        /// 自定义要素类
        /// </summary>
        public class Features
        {
            /// <summary>
            /// 13年宗地代码
            /// </summary>
            public string ZDDMold;
            /// <summary>
            /// 17年宗地代码
            /// </summary>
            public string ZDDMnew;
            /// <summary>
            /// 13年承包方姓名
            /// </summary>
            public string MCold;
            /// <summary>
            /// 17年承包方姓名
            /// </summary>
            public string MCnew;
            /// <summary>
            /// 13年地块面积
            /// </summary>
            public string shapeAreaOld;
            /// <summary>
            /// 17年地块面积
            /// </summary>
            public string shapeAreaNew;
            /// <summary>
            /// 地理要素
            /// </summary>
            public Feature feature;
            /// <summary>
            /// 旧的地块编码
            /// </summary>
            public string DKBMold;
            /// <summary>
            /// 发包方编码，用于定位Excel文件
            /// </summary>
            public string FBFBM;
            /// <summary>
            /// 地块预编码
            /// </summary>
            public string DKYBM;
            /// <summary>
            /// 13年地块预编码
            /// </summary>
            public string DKYBMold;
            /// <summary>
            /// 记录特殊情况
            /// </summary>
            public string Memo = "";
            /// <summary>
            /// 地块类别
            /// </summary>
            public string DKLB;
            public object clone()
            {
                return MemberwiseClone();
            }
        }
        public class Bsm
        {
            public string code;
            public string location;
        }

        private string shiftFeatureString(Feature feature, int idx)
        {
            //调用gdal方法传入feature和属性位置的index
            IntPtr pStr = OGR_F_GetFieldAsString(Feature.getCPtr(feature), idx);
            string f = feature.GetFieldAsString(idx);
            //然后调用.net的非托转托的marshal 把指针转换成变量
            byte[] b = Encoding.Default.GetBytes(f);
            string s = Encoding.UTF8.GetString(b);
            return Marshal.PtrToStringAnsi(pStr).Trim().Replace("\r", "").Replace("\n", "");
            //return s;
        }
        [Obsolete]
        internal static string Utf8BytesToString(IntPtr pNativeData)
        {
            if (pNativeData == IntPtr.Zero)
                return null;

            int nMaxLength = Marshal.PtrToStringAuto(pNativeData).Length;
            int length = 0;//循环查找字符串的长度
            for (int i = 0; i < nMaxLength; i++)
            {
                byte[] strbuf1 = new byte[1];
                Marshal.Copy(pNativeData + i, strbuf1, 0, 1);
                if (strbuf1[0] == 0)
                {
                    break;
                }
                length++;
            }

            byte[] strbuf = new byte[length];
            Marshal.Copy(pNativeData, strbuf, 0, length);
            string s = Encoding.UTF8.GetString(strbuf);
            return s;
        }

        /// <summary>
        /// 线程读取shp文件存到List
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <param name="name">文件名</param>
        /// <param name="dt">输出的datatable</param>
        /// <param name="centers">存储中心点坐标</param>
        private bool ReadShp(string path, string name, /* ISpatialReference iSRef, out DataTable dt,*/ out List<Features> centers)
        {

            ////使用GDAL

            string str = Path.Combine(path, name);
            DataSource ds = Ogr.Open(str, 0);

            Layer layer = ds.GetLayerByIndex(0);  //获取图层
            Feature feature;
            FeatureDefn oDefn = layer.GetLayerDefn();
            string codeName = Gdal.GetConfigOption("SHAPE_ENCODING", "");
            List<Features> cts = new List<Features>();
            //IFeatureCursor iCur = pfc.Search(null, false);//插入游标
            int id = oDefn.GetFieldIndex("DKBM");//地块编码
            int id2 = oDefn.GetFieldIndex("CBFDBXM");
            int id3 = oDefn.GetFieldIndex("Shape_Area");
            int id4 = oDefn.GetFieldIndex("DKBM_1");
            int id5 = oDefn.GetFieldIndex("CBFDBXM_1");
            int id6 = oDefn.GetFieldIndex("Shape_Ar_1");
            int id7 = oDefn.GetFieldIndex("FBFBM");
            int id8 = oDefn.GetFieldIndex("DKBM_1");  //旧的OID
            int id9 = oDefn.GetFieldIndex("DKYBM");   //地块预编码
            int id10 = oDefn.GetFieldIndex("DKYBM_1");//13年地块与编码
            int id11 = oDefn.GetFieldIndex("DKLB");//地块类别
            //IFeature feature;
            while ((feature = layer.GetNextFeature()) != null)
            {
                Features c = new Features();

                c.ZDDMnew = shiftFeatureString(feature, id);  //17年宗地代码赋值
                c.MCnew = shiftFeatureString(feature, id2);//17年承包方姓名
                c.shapeAreaNew = shiftFeatureString(feature, id3);//17年地块面积
                c.ZDDMold = shiftFeatureString(feature, id4);//13年宗地代码
                c.MCold = shiftFeatureString(feature, id5);//13年承包方姓名
                c.shapeAreaOld = shiftFeatureString(feature, id6);//13年地块面积
                c.FBFBM = shiftFeatureString(feature, id7);//发包方编码
                c.DKBMold = shiftFeatureString(feature, id8);
                c.DKYBM = shiftFeatureString(feature, id9);
                c.DKYBMold = shiftFeatureString(feature, id10);
                c.DKLB = shiftFeatureString(feature, id11);//地块类别
                c.feature = feature;//要素

                cts.Add(c);
            }

            //Release
            //Marshal.ReleaseComObject(iCur);
            //Marshal.ReleaseComObject(feature);
            centers = cts;
            return true;
        }
        /// <summary>
        /// 线程读取原始shp文件存到List
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <param name="name">文件名</param>
        /// <param name="dt">输出的datatable</param>
        /// <param name="centers">存储中心点坐标</param>
        private bool ReadShpOrigin(string path, string name, /* ISpatialReference iSRef, out DataTable dt,*/ out List<Features> centers)
        {
            //使用GDAL
            string str = Path.Combine(path, name);
            DataSource ds = Ogr.Open(str, 0);

            Layer layer = ds.GetLayerByIndex(0);  //获取图层
            Feature feature;
            FeatureDefn oDefn = layer.GetLayerDefn();

            string codeName = Gdal.GetConfigOption("SHAPE_ENCODING", "");
            List<Features> cts = new List<Features>();
            //IFeatureCursor iCur = pfc.Search(null, false);//插入游标
            int id = oDefn.GetFieldIndex("DKBM");//地块编码
            int id2 = oDefn.GetFieldIndex("CBFDBXM");
            int id3 = oDefn.GetFieldIndex("Shape_Area");
            //int id4 = pfc.FindField("DKBM_1");
            //int id5 = pfc.FindField("CBFDBXM_1");
            //int id6 = pfc.FindField("Shape_Ar_1");
            int id7 = oDefn.GetFieldIndex("FBFBM");
            //int id8 = oDefn.GetFieldIndex("FID");  //FID
            int id9 = oDefn.GetFieldIndex("DKYBM");   //地块预编码
            int id10 = oDefn.GetFieldIndex("DKLB");
            //IFeature feature;
            while ((feature = layer.GetNextFeature()) != null)
            {
                Features c = new Features();
                c.ZDDMnew = shiftFeatureString(feature, id);  //17年宗地代码赋值
                c.MCnew = shiftFeatureString(feature, id2);//17年承包方姓名
                c.shapeAreaNew = shiftFeatureString(feature, id3);//17年地块面积
                //c.ZDDMold = feature.Value[id4].ToString();//13年宗地代码
                //c.MCold = feature.Value[id5].ToString();//13年承包方姓名
                //c.shapeAreaOld = feature.Value[id6].ToString();//13年地块面积
                c.FBFBM = shiftFeatureString(feature, id7);//发包方编码
                //c.OIDold = shiftFeatureString(feature, id8);
                c.DKYBM = shiftFeatureString(feature, id9);
                c.DKLB = shiftFeatureString(feature, id10);
                c.feature = feature;//要素

                cts.Add(c);
            }

            //Release
            //Marshal.ReleaseComObject(iCur);
            //Marshal.ReleaseComObject(feature);
            centers = cts;
            return true;
        }
        /// <summary>
        /// 自定义openfiledialog类
        /// </summary>
        /// <param name="title">窗体名称</param>
        /// <param name="filter">关键字过滤</param>
        /// <returns></returns>
        private OpenFileDialog myOpen(string title, string filter, bool multiselect = true)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = filter;
            open.Title = title;
            open.RestoreDirectory = false;
            open.Multiselect = multiselect;
            return open;
        }
        /// <summary>
        /// 线程读取Excel文件
        /// </summary>
        /// <param name="filePath">文件完整路径</param>
        /// <param name="workbook">输出第一张表workbook</param>
        /// <param name="dt">输出第二张表DataTable</param>
        private void ReadExcel(string filePath, out IWorkbook workbook, out DataTable dt, out System.IO.FileStream fs)
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
            string fileType = System.IO.Path.GetExtension(filePath);
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
        private void TrimDataTable(DataTable dt)
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
        /// 截去字符串中的各种符号
        /// </summary>
        /// <param name="FieldName"></param>
        /// <returns></returns>
        private string TrimFieldName(string FieldName)
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
        /// 处理DataTable字段名
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="start">查找的第一行第一个字段名称</param>
        /// <param name="headCount">表头字段的行数</param>
        private void EditFieldsName(DataTable dt, string start, int headCount)
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
        /// <summary>
        /// 判断点是否在几何图形上
        /// </summary>
        /// <param name="pGeometry">传入的几何图形</param>
        /// <param name="pPoint">待判断的点</param>
        /// <param name="dTolerance">容差</param>
        /// <returns>是否</returns>

        private static DataSet ds;
        private DataTable dtAll;
        public static string Error = "";
        public static FileStream[] fileStream;
        public static IWorkbook[] workbook;
        public static ISheet[] worksheet;
        public static DataTable[] dt;
        public static List<Bsm> bsm = new List<Bsm>();
        private void button4_Click(object sender, EventArgs e)
        {
            #region 判断文件
            if (!File.Exists(tbOldShp.Text))
            {
                MessageBox.Show("请选择13年shp数据！", "提示");
                return;
            }
            if (!File.Exists(tb13_17shp.Text))
            {
                MessageBox.Show("请选择17年shp数据！", "提示");
                return;
            }
            if (!File.Exists(textBox1.Text))
            {
                MessageBox.Show("请选择标识后的shp数据！", "提示");
                return;
            }
            if (!Directory.Exists(tbNewShp.Text))
            {
                MessageBox.Show("不是正确的路径！", "提示");
                return;
            }
            if (!Directory.Exists(tbCBF.Text))
            {
                MessageBox.Show("请选择承包方调查表！", "提示");
                return;
            }
            #endregion
            Error = string.Empty;
            MemoTotal = string.Empty;
            button4.Enabled = false;
            button7.Enabled = false;
            label2.Text = "正在读取Excel...";
            label2.Refresh();

            GdalConfiguration.ConfigureOgr();
            Gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "YES");
            Gdal.SetConfigOption("SHAPE_ENCODING", "");
            Gdal.AllRegister();
            Ogr.RegisterAll();

            #region 线程读取Excel


            ThreadReadExcel(out fileStream, out workbook, out worksheet, out dt);
            //test 定义标识定位

            for (int i = 0; i < fileStream.Length; i++)
            {
                string code = System.IO.Path.GetFileNameWithoutExtension(fileStream[i].Name);//获取文件名
                code = code.Substring(6, code.Length - 6);
                string address = worksheet[i].GetRow(6).GetCell(7).ToString();
                address = address.Split('区')[1].Split('组')[0] + "组";
                Bsm t = new Bsm();
                t.code = code;
                t.location = address;
                bsm.Add(t);
            }
            #endregion

            #region 读取旧台账
            label2.Text = "正在读取旧台账...";
            label2.Refresh();
            List<string> fileList = new List<string>();
            fileList = GetFile(tbNewShp.Text, fileList, null);
            List<string> fileNeed = new List<string>();
            fileNeed = fileList.Where(p => p.EndsWith("农村土地家庭承包农户基本情况信息表(一).xls")).Select(p => p).ToList();

            //读取旧台账
            ds = new DataSet();
            dtAll = new DataTable();
            //多线程处理Test
            ManualResetEvent finish = new ManualResetEvent(false);
            int maxThreadCount = fileNeed.Count;
            foreach (string file in fileNeed)
            {
                object[] pobjs = new object[3];
                pobjs[0] = file;
                pobjs[1] = ds;
                pobjs[2] = fileNeed.Count;
                //ThreadPool.SetMaxThreads(1000, 1000);

                ThreadPool.QueueUserWorkItem(new WaitCallback(ThreadReading), pobjs);
                if (Interlocked.Decrement(ref maxThreadCount) == 0)
                    finish.Set();
                //ThreadReading(file, ds);

            }
            finish.WaitOne();
            while (ds.Tables.Count < fileNeed.Count)
                Thread.Sleep(100);
            dtAll = GetAllDataTable(ds);
            #endregion

            #region 读取13年和17年shp数据
            label2.Text = "正在读取shp数据...";
            label2.Refresh();
            List<Features> shape13List = new List<Features>();
            if (!ReadShpOrigin(path13, name13, out shape13List))
            {
                button4.Enabled = true;
                button7.Enabled = true;
                return;
            }

            List<Features> shape17List = new List<Features>();
            if (!ReadShpOrigin(path17, name17, out shape17List))
            {
                button7.Enabled = button4.Enabled = true;
                return;
            }


            ////TES GPTOOLS
            //IdentityAnalysisTools(tb13_17shp.Text, tbOldShp.Text);
            #endregion
            #region 读取17年被13年标识后的shp数据
            //label2.Text = "正在读取shp数据...";
            List<Features> shapeDataList = new List<Features>();
            if (!ReadShp(oldPath, oldName, out shapeDataList))
            {
                button4.Enabled = button7.Enabled = true;
                return;
            }

            #endregion

            //处理数据
            label2.Text = "正在筛选数据...";
            label2.Refresh();

            #region 筛选出选中的Excel的文件对应的空间要素信息
            //shp1
            List<Features> chosen = getChosen(shapeDataList, fileStream);//筛选标识shp
            List<Features> chosen13 = getChosen(shape13List, fileStream);//筛选13年shp
            List<Features> chosen17 = getChosen(shape17List, fileStream);//筛选17年shp
            ////shpDeath
            //List<Features> chosenDeath = getChosen(shapeDataDeath, fileStream);
            progressBar1.Maximum = chosen13.Count + 1;
            progressBar1.Minimum = 0;
            progressBar1.Value = 0;
            int pValue = 1;
            //13年DKYBM有，17年DKYBM没有的地块判断为灭失，注意加上FBFBM来判断
            List<Features> chosenDeath = new List<Features>();
            List<Features> diff = new List<Features>();
            foreach (var fe in chosen13)
            {
                progressBar1.Value = pValue;
                pValue++;
                string fbf1 = fe.FBFBM.Substring(6, fe.FBFBM.Length - 6);
                //先通过宗地代码查找13年有而17年没有的地块
                List<Features> tmp = chosen17.Where(p => p.FBFBM.EndsWith(fbf1) &&
                p.DKYBM == fe.DKYBM).Select(p => p).ToList();
                //如果宗地代码没有找到，从标识中查找zddmold是否有
                if (tmp.Count == 0)
                {
                    List<Features> tmpchk = chosen.Where(p => p.FBFBM.EndsWith(fbf1) &&
                      p.DKYBMold == fe.DKYBM).Select(p => p).ToList();
                    if (tmpchk.Count == 0)
                    {
                        //从dt中查询对应人名，并返回户主名

                        int idx = getIndex(fbf1, fileStream);
                        //集体的情况需要转换一下McNew字段，将原“集体”变成“XXXX组”的17年版集体标识
                        if (fe.MCnew == "集体")
                        {
                            string dz1 = worksheet[idx].GetRow(6).GetCell(7).ToString();
                            string[] sp = dz1.Split('区')[1].Split('组');
                            fe.MCnew = sp[0] + "组";
                        }
                        DataRow[] row = dt[idx].Select("CBFMC='" + fe.MCnew + "' or XM='" + fe.MCnew + "'");//

                        if (row.Length > 0)
                        {
                            fe.MCnew = row[0]["CBFMC"].ToString();//更换承包方姓名，归为一户.
                            fe.Memo = "灭失";
                            chosenDeath.Add(fe);
                        }
                        else
                        {
                            List<Bsm> b = bsm.Where(p => fe.FBFBM.Contains(p.code)).Select(p => p).ToList();
                            string t = "发包方: " + b[0].location + " 13年shp中的地块预编码: " + fe.DKYBM +
                                " 无法准确在17年表格数据中定位，需人工检查（疑似灭失）";
                            if (!Error.Contains(t))
                                Error += t + "\n";
                        }

                    }
                }


                //筛选不同的部分
                List<Features> d = chosen17.Where(p => getZDDM(p.ZDDMnew) == getZDDM(fe.ZDDMnew) &&
                p.MCnew != fe.MCnew && fe.MCnew != "集体").Select(p => p).ToList();
                if (d.Count == 1)
                {
                    Features F = new Features();
                    F = d[0].clone() as Features;
                    F.MCold = fe.MCnew;
                    F.ZDDMold = fe.ZDDMnew;
                    F.DKYBMold = fe.DKYBM;
                    F.shapeAreaOld = fe.shapeAreaNew;
                    diff.Add(F);
                }


            }
            #endregion
            #region 筛选数据
            //选出地块和权利人都没变，但是宗地代码变化的地块
            List<Features> errDK = getErrorDKbyZDDM(chosen);
            if (errDK.Count > 0)
            {
                foreach (var err in errDK)
                {
                    Error += string.Format("发包方编码：{0}，权利人：{1}，地块预编码：{2}的地块没有沿用13年地块编码\n",
                        err.FBFBM, err.MCnew, err.DKYBM);
                }
            }
            List<Features> filtList = new List<Features>();
            //将重点属性完全一样的去掉，留下有变化的部分（此处要排除因标识后产生的OID为0的部分）
            //shp1
            filtList = chosen.Where(p => p.MCnew != p.MCold || /*p.shapeAreaNew != p.shapeAreaOld ||*/
            getZDDM(p.ZDDMold) != getZDDM(p.ZDDMnew) ||
            p.DKYBM != p.DKYBMold).Select(p => p).ToList();

            List<Features> NEW = new List<Features>();
            foreach (var fe in chosen17)
            {
                List<Features> neo = chosen13.Where(p => p.DKYBM == fe.DKYBM
                  && getZDDM(p.FBFBM) == getZDDM(fe.FBFBM)).Select(p => p).ToList();
                if (neo.Count == 0)
                {
                    if (!Funcs.isInside(fe.feature, chosen13))
                    {
                        fe.Memo = "新增";
                        NEW.Add(fe);
                    }

                }
            }
            filtList = filtList.Union(NEW).ToList();
            //处理分地，换地的情况。排除新增地块，排除集体发包

            List<Features> fd = filtList.Where(p => p.MCold != "集体" && p.Memo != "新增" &&
            p.DKBMold != "").Select(p => p).ToList();
            List<Features> fdsub = new List<Features>();
            //排除掉灭失地块
            foreach (var f in fd)
            {
                List<Features> fdsubTmp = chosenDeath.Where(p => p.ZDDMnew == f.ZDDMnew).Select(p => p).ToList();
                if (fdsubTmp.Count == 0)
                    fdsub.Add(f);
            }
            //test

            List<Features> fdSum = new List<Features>();
            //先循环处确权错误的部分，将旧地块做标注，为以后分地的情况作考虑
            foreach (var f in fdsub)
            {
                if (f.Memo == "新增") continue;
                if (f.MCnew == f.MCold)
                    continue;
                //List<Features> fdcheck = fdsub.Where(p => getZDDM(p.ZDDMnew) == getZDDM(f.ZDDMnew)).Select(p => p).ToList();
                //if (fdcheck.Count > 1)
                //    continue;
                string fbf1 = f.FBFBM.Substring(6, f.FBFBM.Length - 6);
                int idx = getIndex(fbf1, fileStream);
                DataRow[] rowFD = dtAll.Select("户主编号 like '" + fbf1 + "%' and "
                + "家庭成员姓名 ='" + f.MCold + "'");
                //if (rowFD.Length == 1)
                //{
                //反找旧台账对应整户
                //DataTable dtHH = dtAll.AsEnumerable().Where(p => p.Field<string>("户主编号") ==
                //rowFD[0]["户主编号"].ToString()).Select(p => p).CopyToDataTable();
                //DataRow[] row = dtHH.Select("家庭成员姓名 = '" + f.MCnew + "'");
                //if (row.Length == 0)
                //{
                string fbfmc = getKeyJT(worksheet[idx].GetRow(6).GetCell(7).ToString());

                if (getZDDM(f.ZDDMnew) == getZDDM(f.ZDDMold))
                {
                    List<Features> fqqcw = diff.Where(p => p.MCnew == f.MCold && p.MCold == f.MCnew).Select(p => p).ToList();
                    if (fqqcw.Count == 0)
                    {
                        List<Features> check = chosen.Where(p => p.ZDDMnew == f.ZDDMnew && p.MCnew == p.MCold).Select(p => p).ToList();

                        if (check.Count == 0)
                        {
                            f.Memo = "确权错误";
                            //将旧的chosen13memo修改
                            List<Features> memo13 = chosen13.Where(p => p.ZDDMnew == f.ZDDMold).Select(p => p).ToList();
                            int index = chosen13.FindIndex(p => p.ZDDMnew == f.ZDDMold);
                            chosen13[index].Memo = "确权错";
                            chosen13[index].MCold = f.MCnew;
                            //统一修改确权错误的旧权利人
                            DataRow[] r = dt[idx].Select("XM='" + f.MCold + "'");
                            string hz = "";
                            if (r.Length > 0)
                                hz = r[0]["CBFMC"].ToString();
                            if (hz != f.MCnew && hz != "")
                                f.MCold = hz;

                            ////将旧的chosen13memo修改
                            //List<Features> memo13 = chosen13.Where(p => p.ZDDMnew == f.ZDDMold).Select(p => p).ToList();

                            //if (memo13.Count > 0)
                            //{
                            //    List<Features> memo17 = chosen17.Where(p => p.DKYBM == memo13[0].DKYBM).Select(p => p).ToList();
                            //    if (memo17.Count > 0)
                            //    {
                            //        int index = chosen13.FindIndex(p => p.ZDDMnew == f.ZDDMold);
                            //        chosen13[index].Memo = "确权错";
                            //        chosen13[index].MCold = f.MCnew;
                            //    }

                            //}



                        }

                    }
                }
                //}
                // }
            }
            //List<Features> test = chosen13.Where(p => p.Memo == "确权错").Select(p => p).ToList();
            //第二次循环全部处理
            foreach (var f in fdsub)
            {
                if (f.MCnew == f.MCold || f.Memo == "新增")
                    continue;
                List<Features> fdcheck = fdsub.Where(p => getZDDM(p.ZDDMnew) == getZDDM(f.ZDDMnew)).Select(p => p).ToList();
                if (fdcheck.Count > 1)
                    continue;
                string fbf1 = f.FBFBM.Substring(6, f.FBFBM.Length - 6);
                int idx = getIndex(fbf1, fileStream);
                DataRow[] rowFD = dtAll.Select("户主编号 like '" + fbf1 + "%' and "
                + "家庭成员姓名 ='" + f.MCold + "'");
                if (rowFD.Length == 1)
                {
                    //反找旧台账对应整户
                    DataTable dtHH = dtAll.AsEnumerable().Where(p => p.Field<string>("户主编号") ==
                    rowFD[0]["户主编号"].ToString()).Select(p => p).CopyToDataTable();
                    DataRow[] row = dtHH.Select("家庭成员姓名 = '" + f.MCnew + "'");
                    if (row.Length > 0)
                    {
                        //旧台账承包方姓名在新表中查询，若有则是带地分户，若无则是承包方变更
                        DataRow[] dtfd = dt[idx].Select("CBFMC = '" + f.MCold + "'");
                        if (dtfd.Length > 0)
                            f.Memo = "带地分户";
                    }
                    else
                    {
                        string fbfmc = getKeyJT(worksheet[idx].GetRow(6).GetCell(7).ToString());

                        if (getZDDM(f.ZDDMnew) != getZDDM(f.ZDDMold))
                        {
                            List<Features> ck = chosen13.Where(p => getZDDM(p.ZDDMnew) == getZDDM(f.ZDDMnew)).ToList();
                            if (ck.Count > 0) continue;
                            List<Features> multi = chosen13.Where(p => p.ZDDMnew == f.ZDDMold).ToList(); ;
                            bool isIn = false;
                            foreach (var fe in multi)
                            {
                                if (Funcs.isInside(f.feature, fe.feature))
                                    isIn = true;
                            }
                            if (isIn)
                            {
                                if (multi[0].Memo == "确权错")
                                {
                                    List<Features> chkqq = chosen17.Where(p => p.DKYBM == f.DKYBMold).Select(p => p).ToList();
                                    if (chkqq.Count > 0)
                                        f.MCold = multi[0].MCold;
                                }
                                if (getAreaM(f.feature, 2) != "0")
                                    f.Memo = "分地";//先备注分地，后面还需进一步拆分情况
                            }
                            //分析换地情况
                            //List<Features> multi = fdsub.Where(p => getZDDM(p.ZDDMold) == getZDDM(f.ZDDMold)).Select(p => p).ToList();
                            //if (multi.Count > 1)
                            //{
                            //    List<Features> multiRev = chosen13.Where(p => getZDDM(p.ZDDMnew) == getZDDM(f.ZDDMnew)).Select(p => p).ToList();
                            //    if (multiRev.Count == 0)
                            //    {
                            //        List<Features> c13 = chosen13.Where(p => p.ZDDMnew == f.ZDDMold).Select(p => p).ToList();
                            //        List<Features> c17 = chosen17.Where(p => p.ZDDMnew == f.ZDDMnew).Select(p => p).ToList();
                            //        if (c13.Count > 0)
                            //        {
                            //            if (Funcs.isInside(c17[0].feature, c13[0].feature))
                            //            {
                            //                if (c13[0].Memo == "确权错")
                            //                {
                            //                    List<Features> chkqq = chosen17.Where(p => p.DKYBM == f.DKYBMold).Select(p => p).ToList();
                            //                    if (chkqq.Count > 0)
                            //                        f.MCold = c13[0].MCold;
                            //                }

                            //                f.Memo = "分地";//先备注分地，后面还需进一步拆分情况
                            //            }

                            //        }

                            //    }


                            else
                            {
                                //换地或者是确权错误
                                List<Features> fhd = fdsub.Where(p => p.MCnew == f.MCold && p.MCold == f.MCnew
                                && (getZDDM(p.ZDDMnew) != getZDDM(f.ZDDMold) ||
                                getZDDM(p.ZDDMold) != getZDDM(f.ZDDMnew))).Select(p => p).ToList();

                                if (fhd.Count > 0)
                                {
                                    f.Memo = "换地";
                                    //fhd[0].Memo = "换地出";
                                }
                            }
                        }
                    }
                    fdSum.Add(f);
                }

            }
            List<Features> DDFCout = fdSum.Where(p => p.Memo == "带地分户").Select(p => p).ToList();
            foreach (var f in DDFCout)
            {
                Features F = new Features();
                F = f.clone() as Features;
                F.MCnew = f.MCold;
                F.MCold = f.MCnew;
                F.Memo = "带地分出";
                fdSum.Add(F);
            }
            List<Features> QQCWout = fdsub.Where(p => p.Memo == "确权错误").Select(p => p).ToList();
            foreach (var f in QQCWout)
            {
                //List<Features> chk = chosen17.Where(p => p.MCnew == f.MCold &&
                //  getZDDM(p.FBFBM) == getZDDM(f.FBFBM)).Select(p => p).ToList();
                //if (chk.Count > 0)
                //{
                Features F = new Features();
                F = f.clone() as Features;
                F.MCnew = f.MCold; F.MCold = f.MCnew;
                F.Memo = "确权错误出";
                fdSum.Add(F);
                //}
            }
            List<Features> FDout = fdSum.Where(p => p.Memo == "分地").Select(p => p).ToList();
            foreach (var f in FDout)
            {
                Features F = new Features();
                F = f.clone() as Features;
                //List<Features> chkqqcw = chosen17.Where(p => p.MCnew == f.MCold &&
                //  getZDDM(p.FBFBM) == getZDDM(f.FBFBM)).Select(p => p).ToList();
                //if (chkqqcw.Count > 0)
                //{
                F.MCnew = f.MCold;
                F.MCold = f.MCnew;
                //}
                F.Memo = "分地出";
                fdSum.Add(F);
            }
            List<Features> HDout = fdsub.Where(p => p.Memo == "换地").Select(p => p).ToList();
            foreach (var f in HDout)
            {
                Features F = new Features();
                F = f.clone() as Features;
                F.MCnew = f.MCold; F.MCold = f.MCnew;
                F.Memo = "换地出";
                fdSum.Add(F);
            }

            //将两组shp筛选后的数据合并
            filtList = filtList.Union(chosenDeath).Union(fdSum).ToList();
            //查询17年表对应户主，修改
            foreach (var fe in filtList)
            {
                if (fe.Memo == "确权错误")
                {
                    int idx = getIndex(getZDDM(fe.FBFBM), fileStream);
                    DataRow[] dtRow = dt[idx].Select("XM = '" + fe.MCold + "'");
                    if (dtRow.Length > 0)
                    {
                        if (dtRow[0]["CBFMC"].ToString() != fe.MCnew)
                            fe.MCold = dtRow[0]["CBFMC"].ToString();
                    }
                }


            }
            //按户分组
            //List<Features> order = filtList.OrderBy(p => p.FBFMB).Select(p => p).ToList();
            List<IGrouping<string, Features>> filtGroup = filtList.GroupBy(g => g.MCnew).Select(a => a).ToList();

            #endregion
            //筛选出ZDDM不对的情况
            progressBar1.Visible = true;
            progressBar1.Maximum = filtGroup.Count + 1;
            progressBar1.Minimum = 0;
            pValue = 1;
            foreach (var f in filtGroup)
            {
                if (f.ElementAt(0).ZDDMnew.Length != 19)
                {
                    List<Bsm> b = bsm.Where(p => f.ElementAt(0).FBFBM.Contains(p.code)).Select(p => p).ToList();
                    Error += "发包方: " + b[0].location + " 承包方姓名为: " +
                        f.ElementAt(0).MCnew + " 宗地代码为空或者不是19位\n";
                }

            }

            #region 处理数据
            //处理数据shp1
            foreach (var F in filtGroup)
            {
                progressBar1.Value = pValue;
                pValue++;
                //先从旧台账查询户主信息
                if (F.ElementAt(0).ZDDMnew.Length != 19)
                    continue;
                int length = F.ElementAt(0).ZDDMnew.Length;
                string dm = F.ElementAt(0).ZDDMnew.Substring(6, length - 6);
                string hzbh = dm.Substring(0, dm.Length - 5);
                DataRow[] rowF = dtAll.AsEnumerable().Where(r => r.Field<string>("家庭成员姓名") ==
                F.Key.Trim() && r.Field<string>("户主编号").StartsWith(hzbh) &&
                r.Field<string>("与承包方代表(户主)关系") == "户主").Select(r => r).ToArray();
                //List<Features> dtF = chosen13.Where(p => p.MCnew.Trim() == F.Key.Trim() && p.FBFBM.Contains(hzbh)).Select(p => p).ToList();

                string Memo = "";
                List<Bsm> b = bsm.Where(p => F.ElementAt(0).FBFBM.Contains(p.code)).Select(p => p).ToList();

                string str = b[0].location;
                //int idx = Convert.ToInt32(str.Substring(str.Length - 1, 1));
                List<Features> sub = F.ToList();
                //筛选新增地块，需要将unionDeath合并进来的部分去掉
                //List<Features> subSearchNewBlocksN = sub.Where(p => p.DKBMold == "" &&
                //p.Memo != "灭失" && p.Memo != "带地分户" && p.Memo != "分地" &&
                //p.Memo != "带地分出" && p.Memo != "分地出" && p.Memo != "换地" &&
                //p.Memo != "换地出" && p.Memo != "确权错误" && p.Memo != "确权错误出").Select(p => p).ToList();

                //处理新增地块
                List<Features> subSearchNewBlocks = new List<Features>();
                subSearchNewBlocks = sub.Where(p => p.Memo == "新增").Select(p => p).ToList();
                //foreach (var fe in subSearchNewBlocksN)
                //{
                //    List<Features> tmp = chosen.Where(p => p.FBFBM == fe.FBFBM && p.DKYBM == fe.DKYBM).Select(p => p).ToList();
                //    if (tmp.Count == 1)
                //        subSearchNewBlocks = subSearchNewBlocks.Union(tmp).ToList();
                //}

                //集体发包情况
                string dz = worksheet[0].GetRow(6).GetCell(7).ToString();
                string fbf = getKeyJT(dz);
                List<Features> subSearchJT = sub.Where(p => p.MCold == "集体" &&
                !p.MCnew.StartsWith(fbf) && p.Memo != "灭失" && p.Memo != "带地分户" &&
                p.Memo != "分地" && p.Memo != "带地分出" && p.Memo != "分地出"
                && p.Memo != "换地" && p.Memo != "换地出" && p.Memo != "确权错误" &&
                p.Memo != "确权错误出" && p.Memo != "新增").Select(p => p).ToList();//集体发包

                List<Features> subSearchDeath = sub.Where(p => p.Memo == "灭失").Select(p => p).ToList();//筛选灭失地块 
                List<Features> subSearchDDFH = sub.Where(p => p.Memo == "带地分户").Select(p => p).ToList();//带地分户
                List<Features> subSearchDDFC = sub.Where(p => p.Memo == "带地分出").Select(p => p).ToList();//带地分出
                List<Features> subsearchFD = sub.Where(p => p.Memo == "分地").Select(p => p).ToList();
                List<Features> subsearchFDC = sub.Where(p => p.Memo == "分地出").Select(p => p).ToList();
                List<Features> subsearchHD = sub.Where(p => p.Memo == "换地").Select(p => p).ToList();
                List<Features> subsearchHDC = sub.Where(p => p.Memo == "换地出").Select(p => p).ToList();
                List<Features> subsearchQQCW = sub.Where(p => p.Memo == "确权错误").Select(p => p).ToList();
                List<Features> subsearchQQCWC = sub.Where(p => p.Memo == "确权错误出").Select(p => p).ToList();
                //所有记录都为0则进行下一循环
                if (subSearchNewBlocks.Count == 0 && subSearchJT.Count == 0 && subSearchDeath.Count == 0 &&
                        subSearchDDFH.Count == 0 && subSearchDDFC.Count == 0 && subsearchFD.Count == 0 && subsearchFDC.Count == 0 &&
                        subsearchHD.Count == 0 && subsearchHDC.Count == 0 && subsearchQQCW.Count == 0 && subsearchQQCWC.Count == 0)
                    continue;
                bool gh = false;
                string old = "";
                //查找是否有原户主
                DataTable oldHZ = dtAll.Select("户主编号 like '" + b[0].code
                    + "%' and [与承包方代表(户主)关系] = '户主'").CopyToDataTable();

                if (rowF.Length >= 1)//有且匹配户主，为漏登地块
                {
                    Memo = "发包方: " + str + " 承包方姓名: " + F.Key + " ";
                }
                else if (rowF.Length < 1)//没找到户主
                {
                    if (F.Key.Contains(fbf))//集体情况排除
                        Memo = "发包方: " + str + " 承包方姓名: " + F.Key + " ";
                    else
                    {
                        int idx = getIndex(b[0].code, fileStream);
                        DataRow[] row = dt[idx].Select("CBFMC = '" + F.Key + "'");
                        for (int i = 0; i < row.Length; i++)
                        {
                            DataRow[] R = oldHZ.Select("家庭成员姓名 = '" + row[i]["XM"].ToString() + "'");
                            if (R.Length > 0)
                            {
                                gh = true;
                                old = row[i]["XM"].ToString();
                                break;
                            }
                        }
                        if (gh == false)
                            Memo = "发包方: " + str + " 承包方姓名: " + F.Key + " 本户为新增承包方，";
                        else
                        {
                            Memo = "发包方: " + str + " 承包方姓名: " + F.Key + " 本户要求更换承包方代表，从"
                                + old + "变为" + F.Key + "；";
                        }
                    }

                }
                else
                {
                    string t = "发包方: " + str + " 承包方姓名为: " + F.Key + " 在旧台账中找到多个相同姓名，无法准确定位。需人工检查";
                    if (!Error.Contains(t))
                        Error += t + "\n";
                    continue;
                }
                if (subSearchNewBlocks.Count > 0)
                {
                    Memo += "本户因2013年确权漏登增加地块";
                    foreach (var fe in subSearchNewBlocks)
                    {
                        Memo += string.Format(getDKBM(fe.ZDDMnew) + "{0}" + "，", fe.DKLB == "21" ?
                            "（自留地）" : "");
                    }
                    Memo = Memo.Remove(Memo.Length - 1, 1) + "；";
                }
                if (subSearchJT.Count > 0)
                {
                    Memo += "本户地块";
                    foreach (var fe in subSearchJT)
                    {
                        Memo += getDKBM(fe.ZDDMnew) + "，";
                    }
                    Memo = Memo.Remove(Memo.Length - 1, 1) + "因2013年确权登记错误为" + str + "集体，现确权为本户；";
                }
                if (subSearchDeath.Count > 0)
                {
                    Memo += "本户地块";
                    foreach (var fe in subSearchDeath)
                    {
                        Memo += getDKBM(fe.ZDDMnew) + "，";
                    }
                    Memo = Memo.Remove(Memo.Length - 1, 1) + "灭失；";
                }
                if (subSearchDDFH.Count > 0)
                {
                    Memo += "本户由" + judgeName(subSearchDDFH[0].MCold, fbf) + "分户迁出，地块";
                    foreach (var fe in subSearchDDFH)
                    {
                        Memo += getDKBM(fe.ZDDMnew) + "，";
                    }
                    Memo = Memo.Remove(Memo.Length - 1, 1) + "分给本户；";
                }
                if (subSearchDDFC.Count > 0)
                {
                    Memo += "本户分出权利人" + judgeName(subSearchDDFC[0].MCold, fbf) + "，分出地块";
                    foreach (var fe in subSearchDDFC)
                    {
                        Memo += getDKBM(fe.ZDDMnew) + "，";
                    }
                    Memo = Memo.Remove(Memo.Length - 1, 1) + "；";
                }
                if (subsearchFD.Count > 0)
                {
                    foreach (var fe in subsearchFD)
                    {
                        Memo += "本户地块" + getDKBM(fe.ZDDMnew) + "，由权利人"
                            + judgeName(fe.MCold, fbf) + "的" + getDKBM(fe.ZDDMold)
                            + "因确权错误地块分割而来" + "（" + getAreaM(fe.feature, 2)
                            + "亩）；";
                    }
                }
                if (subsearchFDC.Count > 0)
                {
                    foreach (var fe in subsearchFDC)
                    {
                        Memo += "本户地块" + getDKBM(fe.ZDDMold)
                            + "，分割给权利人" + judgeName(fe.MCold, fbf) + "的" + getDKBM(fe.ZDDMnew)
                            + "（" + getAreaM(fe.feature, 2) + "亩）；";
                    }
                }
                if (subsearchHD.Count > 0)
                {
                    List<IGrouping<string, Features>> To = subsearchHD.GroupBy(p => p.MCold).ToList();

                    foreach (var gp in To)
                    {
                        Memo += "本户地块";
                        List<Features> dk = subsearchHD.Where(p => p.MCold == gp.Key).Select(p => p).ToList();
                        foreach (var fe in dk)
                            Memo += getDKBM(fe.ZDDMnew) + "，";
                        Memo = Memo.Remove(Memo.Length - 1, 1) + "与" + judgeName(gp.Key, fbf) + "地块";
                        foreach (var fe in dk)
                            Memo += getDKBM(fe.ZDDMold) + "，";
                        Memo = Memo.Remove(Memo.Length - 1, 1) + "对换；";
                    }

                }
                if (subsearchHDC.Count > 0)
                {
                    List<IGrouping<string, Features>> To = subsearchHDC.GroupBy(p => p.MCold).ToList();

                    foreach (var gp in To)
                    {
                        Memo += "本户地块";
                        List<Features> dk = subsearchHDC.Where(p => p.MCold == gp.Key).Select(p => p).ToList();
                        foreach (var fe in dk)
                            Memo += getDKBM(fe.ZDDMnew) + "，";
                        Memo = Memo.Remove(Memo.Length - 1, 1) + "与" + judgeName(gp.Key, fbf) + "地块";
                        foreach (var fe in dk)
                            Memo += getDKBM(fe.ZDDMold) + "，";
                        Memo = Memo.Remove(Memo.Length - 1, 1) + "对换；";
                    }

                }
                if (subsearchQQCW.Count > 0)
                {
                    List<IGrouping<string, Features>> To = subsearchQQCW.GroupBy(p => p.MCold).ToList();
                    foreach (var gp in To)
                    {
                        if (gp.Key == old) continue;
                        Memo += "原" + judgeName(gp.Key, fbf) + "地块";
                        List<Features> dk = subsearchQQCW.Where(p => p.MCold == gp.Key).Select(p => p).ToList();
                        foreach (var fe in dk)
                            Memo += getDKBM(fe.ZDDMnew) + "，";
                        Memo = F.Key.Contains(fbf) ? Memo.Remove(Memo.Length - 1, 1) + "因确权错误退还集体；"
                            : Memo.Remove(Memo.Length - 1, 1) + "因确权错误，现确为本户；";
                    }
                }
                if (subsearchQQCWC.Count > 0)
                {
                    List<IGrouping<string, Features>> To = subsearchQQCWC.GroupBy(p => p.MCold).ToList();
                    foreach (var gp in To)
                    {
                        Memo += "地块";
                        List<Features> dk = subsearchQQCWC.Where(p => p.MCold == gp.Key).Select(p => p).ToList();
                        foreach (var fe in dk)
                            Memo += getDKBM(fe.ZDDMnew) + "，";
                        Memo = gp.Key.Contains(fbf) ? Memo.Remove(Memo.Length - 1, 1) + "因确权错误退还集体；"
                            : Memo.Remove(Memo.Length - 1, 1) + "因确权错误，现确为" + judgeName(gp.Key, fbf) + "；";
                    }
                }
                if (Memo != "")
                {
                    Memo = Memo.Remove(Memo.Length - 1, 1);//去掉最后一个多余的分号
                    if (Memo.Contains("分户迁出"))
                        Memo = Memo.Replace("本户为新增承包方，", "");
                    MemoTotal += Memo + "\n";
                }
            }

            //if (Error != "")
            //    MemoTotal = Error + "\n" + MemoTotal;
            #endregion
            #region 显示结果
            showResult();
            #endregion
            //MessageBox.Show(MemoTotal);
            //MessageBox.Show(Error);
            button4.Enabled = true;
            button7.Enabled = true;
            progressBar1.Visible = false;
            MessageBox.Show("完成!");
            label2.Text = string.Empty;
        }
        private string oldPath, newPath, oldName, newName, path13, name13, path17, name17;  //存储13年与17年shp数据的路径与文件名
        private string[] cbfFiles;  //存储承包方调查表台账（整村）
        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = myOpen("请选择承包地调查表", "Excel97-2003|*.xls");
            tbCBF.Clear();
            if (open.ShowDialog() == DialogResult.OK)
            {
                foreach (string s in open.FileNames)
                {
                    IntPtr vHandle = _lopen(s, OF_READWRITE | OF_SHARE_DENY_NONE);
                    if (vHandle == HFILE_ERROR)
                    {
                        MessageBox.Show("文件 " + Path.GetFileName(s) +
                            " 已被打开，请先关闭文件。");
                        return;
                    }
                    CloseHandle(vHandle);
                }

                tbCBF.Text = Path.GetDirectoryName(open.FileName);
                cbfFiles = open.FileNames;



            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog open = new FolderBrowserDialog();
            open.Description = "请选择13年旧台账路径";
            //OpenFileDialog open = myOpen("请选择13年旧台账", "Excel97-2003|*.xls", false);
            tbNewShp.Clear();
            if (open.ShowDialog() == DialogResult.OK)
            {
                tbNewShp.Text = open.SelectedPath;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = myOpen("请选择13年shp数据", "ArcGIS Shp文件|*.shp", false);
            tbOldShp.Clear();
            if (open.ShowDialog() == DialogResult.OK)
            {
                IntPtr vHandle = _lopen(open.FileName, OF_READWRITE | OF_SHARE_DENY_NONE);
                if (vHandle == HFILE_ERROR)
                {
                    MessageBox.Show("文件已被打开，请先关闭文件。");
                }
                else
                {
                    path13 = Path.GetDirectoryName(open.FileName);
                    tbOldShp.Text = open.FileName;
                    name13 = Path.GetFileName(open.FileName);
                }
                CloseHandle(vHandle);

            }
        }

        private void ThreadReadExcel(out FileStream[] fileStream, out IWorkbook[] workbook, out ISheet[] worksheet, out DataTable[] dataTable)
        {
            #region 读取承包方调查表
            int count = cbfFiles.Length;
            IWorkbook[] wb = new IWorkbook[count];
            ISheet[] ws = new ISheet[count];
            DataTable[] dt = new DataTable[count];
            System.IO.FileStream[] fs = new System.IO.FileStream[count];
            Thread[] thread = new Thread[count];
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

        /// <summary>
        ///  复制shp文件
        /// </summary>
        /// <param name="originWs">输入要素的工作空间</param>
        /// <param name="pfc">输入的要素集</param>
        /// <param name="copyPath">另存路径</param>
        /// <param name="copyName">另存文件名</param>
        private static void copyShpFilebyFeatures(string path, string name, string copyPath, string copyName)
        {
            // IWorkspaceFactory pwsf = new ShapefileWorkspaceFactoryClass();
            // IWorkspace pws = pwsf.OpenFromFile(path, 0);
            // IFeatureWorkspace pfeatWs = (IFeatureWorkspace)pws;
            // IFeatureClass pfc = pfeatWs.OpenFeatureClass(name);
            // IGeoDataset pGeoDs = (IGeoDataset)pfc;
            // ISpatialReference spartialRef = pGeoDs.SpatialReference;

            // IDataset inputDataset = (IDataset)pfc;
            // IDatasetName inputDatasetName = (IDatasetName)inputDataset.FullName;

            // IWorkspaceFactory wsf = new ShapefileWorkspaceFactoryClass();
            // IWorkspace ws = null;
            // try
            // {
            //     ws = wsf.OpenFromFile(copyPath, 0);
            // }
            // catch (Exception e)
            // {
            //     Console.WriteLine(e.Message);
            // }
            // IDataset ds = (IDataset)ws;
            // IWorkspaceName wsName = (IWorkspaceName)ds.FullName;
            // IFeatureClassName featClsName = new FeatureClassNameClass();
            // IDatasetName dsName = (IDatasetName)featClsName;
            // dsName.WorkspaceName = wsName;
            // dsName.Name = copyName;

            // //// Use the IFieldChecker interface to make sure all of the field names are valid for a shapefile. 
            // IFieldChecker fieldChecker = new FieldCheckerClass();
            // IFields shapefileFields = null;
            // IEnumFieldError enumFieldError = null;
            // fieldChecker.InputWorkspace = inputDataset.Workspace;
            // fieldChecker.ValidateWorkspace = ws;
            // fieldChecker.Validate(pfc.Fields, out enumFieldError, out shapefileFields);

            // // At this point, reporting/inspecting invalid fields would be useful, but for this example it's omitted.

            // // We also need to retrieve the GeometryDef from the input feature class. 
            // int shapeFieldPosition = pfc.FindField(pfc.ShapeFieldName);
            // IFields inputFields = pfc.Fields;

            // IField shapeField = inputFields.get_Field(shapeFieldPosition);
            // IGeometryDef geometryDef = shapeField.GeometryDef;

            // IGeometryDef pGeometryDef = new GeometryDef();
            // IGeometryDefEdit pGeometryDefEdit = pGeometryDef as IGeometryDefEdit;
            // pGeometryDefEdit.GeometryType_2 = esriGeometryType.esriGeometryPolyline;
            // pGeometryDefEdit.SpatialReference_2 = spartialRef;

            // // Get the layer's selection set. 
            // ISelectionSet selectionSet = pfc.Select(null, esriSelectionType.esriSelectionTypeIDSet,
            //     esriSelectionOption.esriSelectionOptionNormal, null);

            // // Now we can create a feature data converter. 
            // IFeatureDataConverter2 featureDataConverter2 = new FeatureDataConverterClass();
            // IEnumInvalidObject enumInvalidObject = featureDataConverter2.ConvertFeatureClass(inputDatasetName, null,
            //selectionSet, null, featClsName, pGeometryDef, shapefileFields, "", 1000, 0);

            // // Again, checking for invalid objects would be useful at this point...
            // ds = null;
            // ws = null;
            // wsf = null;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label2.Text = string.Empty;
            progressBar1.Visible = false;
        }
        public static string MemoTotal = "";
        /// <summary>
        /// 复制shp文件
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="sourceName"></param>
        /// <param name="desPath"></param>
        /// <param name="desName"></param>
        private static void copyShpFile(string sourcePath, string sourceName, string desPath, string desName)
        {
            string[] files = Directory.GetFiles(sourcePath,
                System.IO.Path.GetFileNameWithoutExtension(sourceName) + ".*");
            for (int i = 0; i < files.Length; i++)
            {
                string ext = System.IO.Path.GetExtension(files[i]); // 扩展名
                string desFull = System.IO.Path.Combine(desPath, desName) + ext;
                File.Copy(files[i], desFull);
            }
        }

        /// <summary>
        /// 清空指定的文件夹，但不删除文件夹
        /// </summary>
        /// <param name="dir">需要清空的路径</param>
        private void DeleteFolder(string dir)
        {
            foreach (string d in Directory.GetFileSystemEntries(dir))
            {
                if (File.Exists(d))
                {
                    try
                    {
                        FileInfo fi = new FileInfo(d);
                        if (fi.Attributes.ToString().IndexOf("ReadOnly") != -1)
                            fi.Attributes = FileAttributes.Normal;
                        File.Delete(d);//直接删除其中的文件 
                    }
                    catch
                    {

                    }
                }
                else
                {
                    try
                    {
                        DirectoryInfo d1 = new DirectoryInfo(d);
                        if (d1.GetFiles().Length != 0)
                        {
                            DeleteFolder(d1.FullName);////递归删除子文件夹
                        }
                        Directory.Delete(d);
                    }
                    catch
                    {

                    }
                }
            }

        }

        /// <summary>  
        /// 获取路径下所有文件以及子文件夹中文件  
        /// </summary>  
        /// <param name="path">全路径根目录</param>  
        /// <param name="FileList">存放所有文件的全路径</param>  
        /// <param name="RelativePath"></param>  
        /// <returns></returns>  
        public static List<string> GetFile(string path, List<string> FileList, string RelativePath)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            FileInfo[] fil = dir.GetFiles();
            DirectoryInfo[] dii = dir.GetDirectories();
            foreach (FileInfo f in fil)
            {
                //int size = Convert.ToInt32(f.Length);  
                //long size = f.Length;
                FileList.Add(f.FullName);//添加文件路径到列表中  
            }
            //获取子文件夹内的文件列表，递归遍历  
            foreach (DirectoryInfo d in dii)
            {
                GetFile(d.FullName, FileList, RelativePath);
            }
            return FileList;
        }

        #region 线程处理子函数
        /// <summary>
        /// 读取和处理DataTable到DataSet
        /// </summary>
        /// <param name="file"></param>
        /// <param name="ds"></param>
        private void ThreadReading(object param)
        {
            object[] jobarr = param as object[];
            string file = jobarr[0] as string;

            int done = Convert.ToInt32(jobarr[2]);

            #region 读取和处理DataTable
            //从文件名获取户主编号
            string[] sp = file.ToString().Split('-');
            string hzbh = sp[3].Substring(6, sp[3].Length - 6);

            string filepath = file;
            System.Data.DataTable dt = getExcelByName(filepath);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                dt.Columns[i].ColumnName = dt.Columns[i].ColumnName.Replace("_", "");  //处理字段名中的回车符
                dt.Columns[i].ColumnName = dt.Columns[i].ColumnName.Replace("\n", "");
            }
            int l = 0;
            while (dt.Rows[l][0].ToString().Trim() != "")
            {
                l++;
            }
            for (int i = dt.Rows.Count - 1; i >= l; i--)
            {
                dt.Rows.RemoveAt(i);//删除后面多余数据
            }
            dt.Columns.Add("户主编号", typeof(string));
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dt.Rows[i]["户主编号"] = hzbh;
            }

            //重命名dt
            dt.TableName = hzbh;
            System.Data.DataTable dt2 = dt.Copy();  //不加这一句会出现此datatable已属于另一个dataset的错误。

            //LOCK
            //object lockjob = new object();
            lock (ds)
            {
                //ds = jobarr[1] as DataSet;
                ds.Tables.Add(dt2);
            }
            cur++;  //计数器加1

            #endregion
        }
        #endregion
        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = myOpen("请选择标识后的shp文件", "ArcGIS Shp文件|*.shp", false);
            textBox1.Clear();
            if (open.ShowDialog() == DialogResult.OK)
            {
                IntPtr vHandle = _lopen(open.FileName, OF_READWRITE | OF_SHARE_DENY_NONE);
                if (vHandle == HFILE_ERROR)
                {
                    MessageBox.Show("文件已被打开，请先关闭文件。");
                }
                else
                {
                    oldPath = System.IO.Path.GetDirectoryName(open.FileName);
                    textBox1.Text = open.FileName;
                    oldName = System.IO.Path.GetFileName(open.FileName);
                }
                CloseHandle(vHandle);

            }
        }

        private int cur = 0;
        #region 读取EXCEL_通用，通过excel表名获取datatable
        /// <summary>
        /// 读取EXCEL到System.Data.DataTable
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="sheetName">表名，默认为Sheet1</param>
        /// <param name="start">读取表格的起始行数，默认为1</param>
        /// <returns></returns>
        public static System.Data.DataTable getExcelByName(string filePath)
        {

            //bool hasTitle = false;
            string fileType = System.IO.Path.GetExtension(filePath);

            using (DataSet ds = new DataSet())
            {
                //string strCon = string.Format("Provider=Microsoft.{0}.OLEDB.{1}.0;" +
                //                    "Extended Properties=\"Excel {2}.0;HDR=YES;IMEX=1;\";" +
                //                    "data source={4};",
                //                    (fileType == ".xls" ? "Jet" : "ACE"), (fileType == ".xls" ? 4 : 12), (fileType == ".xls" ? 8 : 12), (hasTitle ? "Yes" : "NO"), filePath);
                string strCon = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1'";
                string strCom = "SELECT * FROM [Sheet1$A3:I17]";
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

        #endregion
        private void button7_Click(object sender, EventArgs e)
        {
            showResult();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = myOpen("请选择17年shp数据", "ArcGIS Shp文件|*.shp", false);
            tb13_17shp.Clear();
            if (open.ShowDialog() == DialogResult.OK)
            {
                IntPtr vHandle = _lopen(open.FileName, OF_READWRITE | OF_SHARE_DENY_NONE);
                if (vHandle == HFILE_ERROR)
                {
                    MessageBox.Show("文件已被打开，请先关闭文件。");
                }
                else
                {
                    path17 = System.IO.Path.GetDirectoryName(open.FileName);
                    tb13_17shp.Text = open.FileName;
                    name17 = System.IO.Path.GetFileName(open.FileName);
                }
                CloseHandle(vHandle);

            }
        }

        private void tbNewShp_MouseClick(object sender, MouseEventArgs e)
        {
            tbNewShp.Select(0, tbNewShp.Text.Length);
        }

        private void tbOldShp_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.button1_Click(sender, e);
        }

        private void tb13_17shp_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            button5_Click(sender, e);
        }

        private void textBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            button6_Click(sender, e);
        }

        private void tbNewShp_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            button2_Click(sender, e);
        }

        private void tbCBF_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            button3_Click(sender, e);
        }

        /// <summary>
        /// 从Excel表中地址截取查询关键字
        /// </summary>
        /// <param name="dz"></param>
        /// <returns></returns>
        private string getKeyJT(string dz)
        {
            string[] sp = dz.Split('区');
            string[] sp2 = sp[1].Split('村');
            return sp2[0];
        }

        /// <summary>
        /// 根据选择的Excel文件筛选相应的shpList数据
        /// </summary>
        /// <param name="shapeDataList"></param>
        /// <param name="fileStream"></param>
        /// <returns></returns>
        private List<Features> getChosen(List<Features> shapeDataList, System.IO.FileStream[] fileStream)
        {
            progressBar1.Visible = true;
            progressBar1.Maximum = shapeDataList.Count + 1;
            progressBar1.Minimum = 0;
            progressBar1.Value = 0;
            int pValue = 1;
            List<Features> chosen = new List<Features>();
            foreach (var item in fileStream)
            {
                progressBar1.Value = pValue;
                progressBar1.Refresh();
                pValue++;
                string fName = System.IO.Path.GetFileNameWithoutExtension(item.Name);   //获取文件名
                string key = fName.Substring(6, fName.Length - 6);//获取发包方编码关键字
                List<Features> sub = new List<Features>();
                sub = shapeDataList.Where(p => p.FBFBM.EndsWith(key)).Select(p => p).ToList();
                chosen = chosen.Union(sub).ToList();
            }
            return chosen;
        }

        /// <summary>
        /// 根据发包方编码查找对应的文件索引
        /// </summary>
        /// <param name="fbf">发包方编码</param>
        /// <param name="fs">文件流数组</param>
        /// <returns></returns>
        public static int getIndex(string fbf, FileStream[] fs)
        {
            int idx = -1;
            for (int i = 0; i < fs.Length; i++)
            {
                if (fs[i].Name.Contains(fbf))
                    idx = i;
            }
            return idx;
        }
        private void showResult()
        {
            Form2 frm = new Form2();
            frm.dataGridView1.Rows.Clear();
            frm.dataGridView1.RowsDefaultCellStyle.Font = new Font("宋体", 10, FontStyle.Regular);
            //frm.richTextBox1.Text = string.Empty;
            frm.richTextBox1.Text = Error;

            string s = MemoTotal;
            if (Error != "")//有错误的情况下将错误的文字标红
            {

                frm.richTextBox1.Select(0, Error.Length);
                frm.richTextBox1.SelectionColor = Color.Red;
                frm.richTextBox1.Select(0, 0);   //清除选择
                s = MemoTotal.Replace(Error, "");
            }

            string[] sp = s.Split('\n');

            foreach (var str in sp)
            {
                if (str.Length > 3)
                {
                    string[] spBlank = str.Split(' ');
                    if (spBlank.Length == 5)
                    {
                        string[] input = new string[3];
                        input[0] = spBlank[1]; input[1] = spBlank[3]; input[2] = spBlank[4];
                        frm.dataGridView1.Rows.Add(input);
                        frm.dataGridView1.Rows[frm.dataGridView1.RowCount - 1].Cells[2].ToolTipText = "双击复制到剪贴板";
                    }

                }
                //frm.dataGridView1.Sort(frm.dataGridView1.Columns[0], System.ComponentModel.ListSortDirection.Ascending);
            }
            //参数传递
            Funcs.fsPass = fileStream;
            Funcs.dtPass = dt;
            Funcs.bsmPass = bsm;
            frm.Show();
        }
        #region 合并结构相同的DataTable
        /// <summary>
        /// 合并结构相同的DataTable
        /// </summary>
        /// <param name="ds">包含DataTable的DataSet</param>
        /// <returns>合并后的DataTable</returns>
        public static DataTable GetAllDataTable(DataSet ds)
        {
            System.Data.DataTable newDataTable = ds.Tables[0].Clone();                //创建新表 克隆以有表的架构。
            object[] objArray = new object[newDataTable.Columns.Count];   //定义与表列数相同的对象数组 存放表的一行的值。
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                for (int j = 0; j < ds.Tables[i].Rows.Count; j++)
                {
                    ds.Tables[i].Rows[j].ItemArray.CopyTo(objArray, 0);    //将表的一行的值存放数组中。
                    newDataTable.Rows.Add(objArray);                       //将数组的值添加到新表中。
                }
            }
            return newDataTable;                                           //返回新表。
        }
        #endregion

        private string getDKBM(string DKBM)
        {
            if (DKBM.Length < 5)
                return "";
            string s = DKBM.Substring(DKBM.Length - 5, 5);
            return s;
        }
        private string getZDDM(string ZDDM)
        {
            if (ZDDM.Length < 6)
                return "";
            string s = ZDDM.Substring(6, ZDDM.Length - 6);
            return s;
        }
        /// <summary>
        /// 将面积字符（平方米）转换为亩
        /// </summary>
        /// <param name="feature">平方米面积字符串</param>
        /// <param name="dec">保留的小数位数</param>
        /// <returns></returns>
        private string getAreaM(Feature feature, int dec = 1)
        {
            double area = feature.GetGeometryRef().GetArea();
            //double area = Convert.ToDouble(feature);
            area *= 0.0015d;
            area = Math.Round(area, dec);
            return area.ToString();
        }
        private string judgeName(string name, string fbf)
        {
            if (name.Contains(fbf))
                return name;
            else
                return name + "（户）";
        }
        /// <summary>
        /// 查找地块预编码与权利人都没变，而宗地代码变了的地块
        /// </summary>
        /// <param name="chosen13"></param>
        /// <param name="chosen17"></param>
        /// <returns></returns>
        private List<Features> getErrorDKbyZDDM(List<Features> chosen)
        {
            List<Features> result = new List<Features>();
            //foreach (var fe in chosen17)
            //{
            //    List<Features> sel = chosen13.Where(p => p.DKYBM != fe.DKYBM &&
            //      p.MCnew == fe.MCnew && getZDDM(p.ZDDMnew) != getZDDM(fe.ZDDMnew)).Select(p => p).ToList();
            //    if (sel.Count > 0)
            //        result.Add(fe);
            //}
            List<Features> res = chosen.Where(p => p.DKYBM != p.DKYBMold &&
                    p.MCnew == p.MCold && getZDDM(p.ZDDMnew) != getZDDM(p.ZDDMold)).Select(p => p).ToList();
            foreach (var fe in res)
            {
                List<Features> r = chosen.Where(p => getZDDM(p.ZDDMnew) == getZDDM(fe.ZDDMold) &&
                p.MCnew == fe.MCold).Select(p => p).ToList();
                if (r.Count == 0)
                    result.Add(fe);
            }
            return result;
        }
    }
}
