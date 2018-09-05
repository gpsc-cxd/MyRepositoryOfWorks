using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace 承包地13年与17年shp对比
{
    public partial class Form2 : Form
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
        public Form2()
        {
            InitializeComponent();
        }
        //bool isSaved;
        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "Excel97-2003|*.xls";
            save.Title = "保存结果";
            save.RestoreDirectory = false;
            //save.FileName = ".txt";
            if (save.ShowDialog() == DialogResult.OK)
            {
                if (Path.GetExtension(save.FileName).Contains("txt"))
                {
                    //建立txt文本保存信息
                    saveToTXT(save.FileName);
                }
                else if (Path.GetExtension(save.FileName).Contains("rtf"))
                {
                    //保存为rft格式
                    saveToRTF(save.FileName);
                }
                else
                {
                    //保存为excel
                    saveToEXCEL(save.FileName);
                }

                MessageBox.Show("保存成功！");
                //isSaved = true;
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            //isSaved = false;
            fs = Funcs.fsPass;
            datatable = Funcs.dtPass;
            bsm = Funcs.bsmPass;
            if (dataGridView1.Rows.Count == 0)
                checkBox1.Enabled = false;
            else checkBox1.Enabled = true;
            if (fs != null && fs.Length > 0)
                checkBox1.Checked = true;
            else checkBox1.Checked = false;
            button3.Enabled = checkBox1.Checked;

        }

        private void ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells[2].Selected == true)
                    dataGridView1.Rows[i].Cells[2].Style.BackColor = Color.Green;//绿色标记已检查
            }
        }
        private new Point StartPosition;//获取鼠标右键单击的位置
        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                string s = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                s = ClearHtmlTags(s);
                Clipboard.SetText(s);
            }

        }
        private string ClearHtmlTags(object value)
        {
            string str = value.ToString();
            Regex objRegExp = new Regex("<(.|\n)+?>");
            string newStr = objRegExp.Replace(str, "");
            newStr = newStr.Replace("<", "&lt;");
            newStr = newStr.Replace(">", "&gt;");
            return newStr;
        }
        private FileStream[] fs;
        private DataTable[] datatable;
        private List<Form1.Bsm> bsm;
        private void saveToTXT(string path)
        {
            FileStream fa = new FileStream(path, FileMode.Create);
            StreamWriter sw = new StreamWriter(fa);
            sw.Write(richTextBox1.Text.Replace("\n", "\r\n")
                + "\r\n" + Form1.MemoTotal.Replace("\n", "\r\n"));
            sw.Close();
        }
        private void saveToRTF(string path)
        {
            RichTextBox rich = new RichTextBox();
            rich.Font = richTextBox1.Font;
            rich.Text = richTextBox1.Text + "\n" + Form1.MemoTotal;
            rich.Select(0, richTextBox1.Text.Length);
            rich.SelectionColor = Color.Red;
            rich.SaveFile(path);
            //rich.Dispose();
        }
        private void saveToEXCEL(string path)
        {
            IWorkbook workbook; ISheet worksheet;
            workbook = new HSSFWorkbook();
            worksheet = workbook.CreateSheet("对比结果");
            int row = 0;
            if (richTextBox1.Text != "")
            {
                worksheet.CreateRow(0).CreateCell(0).SetCellValue("错误列表");
                string[] sp = richTextBox1.Text.Split('\n');
                row++;
                foreach (var str in sp)
                {
                    if (str.Length > 3)
                    {
                        worksheet.CreateRow(row).CreateCell(0).SetCellValue(str);
                        row++;
                    }
                }
            }
            worksheet.CreateRow(row).CreateCell(0).SetCellValue("发包方");
            worksheet.GetRow(row).CreateCell(1).SetCellValue("承包方姓名");
            worksheet.GetRow(row).CreateCell(2).SetCellValue("调查记事");
            worksheet.GetRow(row).CreateCell(3).SetCellValue("已检查");
            row++;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                worksheet.CreateRow(row).CreateCell(0).SetCellValue(dataGridView1.Rows[i].Cells[0].Value.ToString());
                worksheet.GetRow(row).CreateCell(1).SetCellValue(dataGridView1.Rows[i].Cells[1].Value.ToString());
                worksheet.GetRow(row).CreateCell(2).SetCellValue(dataGridView1.Rows[i].Cells[2].Value.ToString());
                if (dataGridView1.Rows[i].Cells[2].Style.BackColor == Color.Green)
                    worksheet.GetRow(row).CreateCell(3).SetCellValue("True");
                else
                    worksheet.GetRow(row).CreateCell(3).SetCellValue("False");
                row++;
            }
            using (FileStream fsSave = File.OpenWrite(path))
            {
                workbook.Write(fsSave);
                fsSave.Close();
            }
            workbook = null;
            worksheet = null;

        }
        private void openExcel(string path)
        {
            IWorkbook workbook; ISheet worksheet;
            using (FileStream fs = new FileStream(path, FileMode.Open))
            {
                workbook = new HSSFWorkbook(fs);
                worksheet = workbook.GetSheet("对比结果");
                fs.Close();
            }
            if (worksheet == null)
            {
                MessageBox.Show("读取错误，请确定文件的正确性", "错误");
                return;
            }
            dataGridView1.Rows.Clear();
            richTextBox1.Clear();
            int start = 1;

            if (worksheet.GetRow(0).GetCell(0).ToString() == "错误列表")
            {

                while (worksheet.GetRow(start).GetCell(0).ToString() != "发包方")
                {
                    richTextBox1.AppendText(worksheet.GetRow(start).GetCell(0).ToString() + "\n");
                    start++;
                }
                start++;
            }
            for (int i = start; i < worksheet.LastRowNum; i++)
            {
                string fbf = worksheet.GetRow(i).GetCell(0).ToString();
                string cbf = worksheet.GetRow(i).GetCell(1).ToString();
                string dc = worksheet.GetRow(i).GetCell(2).ToString();
                ICell cell = worksheet.GetRow(i).GetCell(3);
                string chk = cell != null ? worksheet.GetRow(i).GetCell(3).ToString() : "False";
                dataGridView1.Rows.Add(fbf, cbf, dc);
                if (chk == "True")
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[2].Style.BackColor = Color.Green;
                start++;
            }
            workbook = null; worksheet = null;
            if (dataGridView1.Rows.Count > 0)
                checkBox1.Enabled = true;
            MessageBox.Show("完成");
        }
        public IContainer cont = new Container();
        public Form3 frm3 = null;
        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //高亮提示
            if (e.Button == MouseButtons.Left && e.RowIndex >= 0 && e.ColumnIndex == 2)
            {
                clearDataGridColor();
                Highlight(e.RowIndex);

            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Title = "打开文件";
            open.Filter = "Excel97-2003|*.xls";
            open.RestoreDirectory = false;
            open.Multiselect = false;
            if (open.ShowDialog() == DialogResult.OK)
            {
                IntPtr vHandle = _lopen(open.FileName, OF_READWRITE | OF_SHARE_DENY_NONE);
                if (vHandle == HFILE_ERROR)
                {
                    MessageBox.Show("文件 " + Path.GetFileName(open.FileName) +
                        " 已被打开，请先关闭文件。");
                    return;
                }
                CloseHandle(vHandle);
                openExcel(open.FileName);
                fs = null;
                checkBox1.Checked = false;
            }
        }

        private void 清除标记ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells[2].Selected == true)
                    dataGridView1.Rows[i].Cells[2].Style.BackColor = Color.White;
            }
        }

        private void contextMenuStrip1_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StartPosition = Cursor.Position;//右键菜单打开同时获取鼠标位置
        }

        private void 清除全部标记ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1.Rows[i].Cells[2].Style.BackColor = Color.White;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                if (fs == null)
                {
                    DialogResult dlgRlt = MessageBox.Show(this, "没有加载需要对比的调查表，是否加载", "提示", MessageBoxButtons.YesNo);
                    if (dlgRlt == DialogResult.No)
                        checkBox1.Checked = false;
                    else
                    {
                        loadDCB();
                        checkBox1.Checked = true;
                    }
                }
            }
            button3.Enabled = checkBox1.Checked;
        }

        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
                checkBox1.Checked = true;
        }

        private void dataGridView1_MouseHover(object sender, EventArgs e)
        {
            this.Activate();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            loadDCB();
        }
        private void loadDCB()
        {
            //加载文件
            OpenFileDialog open = Funcs.myOpenDLG("Excel97-2003|*.xls", "加载承包方调查表");
            if (open.ShowDialog() == DialogResult.Cancel)
            {
                checkBox1.Checked = false;
                return;
            }
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
            fs = null; datatable = null; bsm.Clear();
            string[] files = open.FileNames;
            IWorkbook[] wb; ISheet[] ws;
            Funcs.ThreadReadExcel(files, out fs, out wb, out ws, out datatable);
            //定义标识码
            for (int i = 0; i < fs.Length; i++)
            {
                string code = System.IO.Path.GetFileNameWithoutExtension(fs[i].Name);//获取文件名
                code = code.Substring(6, code.Length - 6);
                string address = ws[i].GetRow(6).GetCell(7).ToString();
                address = address.Split('区')[1].Split('组')[0] + "组";
                Form1.Bsm t = new Form1.Bsm();
                t.code = code;
                t.location = address;
                bsm.Add(t);
            }
        }

        private void splitContainer1_MouseHover(object sender, EventArgs e)
        {
            this.Activate();
        }
        private void markCell(string show, int row)
        {
            string s = dataGridView1.Rows[row].Cells[2].Value.ToString();
            string[] spmark = Funcs.splitCodes(s);
            if (spmark.Length == 0) return;
            //List<string> outStr = new List<string>();
            foreach (var str in spmark)
            {
                if (str.Length != 5) continue;
                if (show.Contains(str))
                {
                    if (!s.Contains("<body bgcolor=\"yellow\">" + str + "</body>"))
                    {
                        if (!s.Contains("<body bgcolor=\"red\">" + str + "</body>"))
                            s = s.Replace(str, "<body bgcolor=\"yellow\">" + str + "</body>");
                        else
                            s = s.Replace("<body bgcolor=\"yellow\">" + str + "</body>", "<body bgcolor=\"red\">" + str + "</body>");
                    }
                }
                else
                {
                    if (!s.Contains("<body bgcolor=\"red\">" + str + "</body>"))
                    {
                        if (!s.Contains("<body bgcolor=\"yellow\">" + str + "</body>"))
                            s = s.Replace(str, "<body bgcolor=\"red\">" + str + "</body>");
                        else
                            s = s.Replace("<body bgcolor=\"yellow\">" + str + "</body>", "<body bgcolor=\"red\">" + str + "</body>");
                    }
                }

            }
            //DOMParser paser = new DOMParser();
            //DOMNode document = paser.parse_text(s);
            dataGridView1.Rows[row].Cells[2].Value = s;
            dataGridView1.Refresh();

        }
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            //自动编号，与数据无关
            Rectangle rectangle = new Rectangle(e.RowBounds.Location.X,
               e.RowBounds.Location.Y,
               dataGridView1.RowHeadersWidth - 4,
               e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics,
                  (e.RowIndex + 1).ToString(),
                   dataGridView1.RowHeadersDefaultCellStyle.Font,
                   rectangle,
                   dataGridView1.RowHeadersDefaultCellStyle.ForeColor,
                   TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection sel = dataGridView1.SelectedRows;

        }
        private void clearDataGridColor()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells[2].Style.BackColor != Color.Green)
                {
                    dataGridView1.Rows[i].Cells[2].Style.BackColor = Color.White;
                }
                string s = dataGridView1.Rows[i].Cells[2].Value.ToString();
                string[] sp = Funcs.splitCodes(s);
                foreach (var str in sp)
                {
                    if (str.Length != 5) continue;
                    if (s.Contains("<body bgcolor=\"yellow\">" + str + "</body>"))
                        s = s.Replace("<body bgcolor=\"yellow\">" + str + "</body>", str);
                    if (s.Contains("<body bgcolor=\"red\">" + str + "</body>"))
                        s = s.Replace("<body bgcolor=\"red\">" + str + "</body>", str);
                }
                dataGridView1.Rows[i].Cells[2].Value = s;
            }
        }
        private void Highlight(int rowIndex)
        {
            string s = dataGridView1.Rows[rowIndex].Cells[2].Value.ToString();
            Funcs.codes = s;
            //if (s.Contains("分割") || s.Contains("分出") ||
            //    s.Contains("迁出") || s.Contains("确权错误") || s.Contains("更换承包方代表"))
            //{

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                string ckn = dataGridView1.Rows[i].Cells[1].Value.ToString();//承包方姓名
                string ckf = dataGridView1.Rows[i].Cells[0].Value.ToString();//发包方
                                                                             //string contain = ckn;
                if (s.Contains("集体"))
                {
                    //是权利人的时候才处理
                    string convFbf = Funcs.convertNumToCha(ckf, false);
                    s = s.Replace("集体", convFbf);
                    //contain = convFbf;
                }
                if (s.Contains(ckn))
                {
                    if (dataGridView1.Rows[i].Cells[2].Style.BackColor != Color.Green)
                        dataGridView1.Rows[i].Cells[2].Style.BackColor = Color.Yellow;
                }
            }
            //}
            //弹框对比
            if (checkBox1.Checked == false)
                return;
            if (fs != null)
            {
                if (fs.Length > 0)
                {
                    string fbf = dataGridView1.Rows[rowIndex].Cells[0].Value.ToString();
                    var sel = bsm.Where(p => p.location.Contains(fbf)).Select(p => p).ToList();
                    if (sel.Count == 0) return;
                    int idx = Form1.getIndex(sel[0].code, fs);
                    DataTable dt = datatable[idx];
                    string cbf = dataGridView1.Rows[rowIndex].Cells[1].Value.ToString();
                    if (cbf.Length > 5)
                    {
                        //承包方为集体德情况
                        cbf = fbf;
                    }
                    DataRow[] row = dt.Select("CBFMC='" + cbf + "'");
                    if (frm3 == null || frm3.IsDisposed)
                    {
                        frm3 = new Form3();

                        frm3.Show(this);
                    }
                    //else
                    //{

                    //    frm3.Refresh();
                    //    frm3.Activate();
                    //}
                    if (row.Length > 0)
                    {
                        string show = string.Format("{0}  {1}\n\n{2}",
                            row[0]["DZ"].ToString(), row[0]["CBFMC"].ToString(), row[0]["DCJS"].ToString());
                        frm3.richTextBox1.Text = show;
                        markCell(show, rowIndex);
                    }
                    else if (row.Length == 0)
                    {
                        frm3.richTextBox1.Text = "原表中未找到姓名一致的对应权利人";
                    }

                    //MessageBox.Show(row[0]["DCJS"].ToString());
                }
            }
        }

        private void dataGridView1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[2].Selected)
                    {
                        clearDataGridColor();
                        Highlight(i);
                    }
                }
            }

        }

        private void Form2_Paint(object sender, PaintEventArgs e)
        {

        }
        
    }
}
