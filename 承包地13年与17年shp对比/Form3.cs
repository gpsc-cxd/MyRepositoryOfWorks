using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 承包地13年与17年shp对比
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }
        protected override void OnLostFocus(EventArgs e)
        {
            base.OnLostFocus(e);

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            string s = Funcs.codes;
            string[] codes = Funcs.splitCodes(s);
            if (codes.Length > 0)
            {
                for (int i = 0; i < codes.Length; i++)
                {
                    if (codes[i].Length != 5) continue;
                    int index = richTextBox1.Text.IndexOf(codes[i]);
                    if (index >= 0)
                        setColor(index, 5, Color.Yellow);
                }
            }
            string[] chk = Funcs.splitCodes(richTextBox1.Text);
            if (chk.Length > 0)
            {
                for(int i = 0; i < chk.Length; i++)
                {
                    if (chk[i].Length != 5) continue;
                    int index = richTextBox1.Text.IndexOf(chk[i]);
                    if (index >= 0)
                    {
                        if (getColor(index, 5) != Color.Yellow)
                            setColor(index, 5, Color.Red);
                    }
                }
            }
        }
        private void setColor(int start, int length, Color color)
        {
            richTextBox1.Select(start, length);
            richTextBox1.SelectionBackColor = color;
            richTextBox1.Select(0, 0);
        }
        private Color getColor(int start,int length)
        {
            richTextBox1.Select(start, length);
            Color color = richTextBox1.SelectionBackColor;
            richTextBox1.Select(0, 0);
            return color;
        }
    }
}
