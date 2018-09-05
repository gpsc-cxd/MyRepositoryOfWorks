using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace 承包地13年与17年shp对比
{
    public class DataGridViewRichTextBoxColumn : DataGridViewTextBoxColumn
    {

        public DataGridViewRichTextBoxColumn() : base()
        {
            CellTemplate = new DataGridViewRichTextBoxCell();
        }
        public override DataGridViewCell CellTemplate
        {
            get
            {
                return base.CellTemplate;
            }

            set
            {
                DataGridViewRichTextBoxCell cell = value as DataGridViewRichTextBoxCell;
                if (value != null && cell == null)
                {
                    throw new InvalidCastException("Value provided for CellTemplate must be of type TEditNumDataGridViewCell or derive from it.");
                }

                base.CellTemplate = value;
            }
        }
    }
    public class DataGridViewRichTextBoxCell : DataGridViewTextBoxCell
    {
        public DataGridViewRichTextBoxCell() : base()
        {

        }
        private static Type defaultEditType = typeof(DataGridViewRTBEditingControl);
        private static Type defaultValueType = typeof(string);
        public override Type EditType
        {
            get
            {
                return base.EditType;
            }
        }
        public override Type ValueType
        {
            get
            {
                return base.ValueType;
            }
        }
        public override void InitializeEditingControl(int rowIndex, object
            initialFormattedValue, DataGridViewCellStyle dataGridViewCellStyle)
        {
            // Set the value of the editing control to the current cell value.
            base.InitializeEditingControl(rowIndex, initialFormattedValue,
                dataGridViewCellStyle);
            DataGridViewRTBEditingControl ctl =
                DataGridView.EditingControl as DataGridViewRTBEditingControl;
            ctl.WordWrap = true;

        }

        static DOMParser parser = new DOMParser();

        public override DataGridViewCellStyle GetInheritedStyle(DataGridViewCellStyle inheritedCellStyle, int rowIndex, bool includeColors)
        {
            return base.GetInheritedStyle(inheritedCellStyle, rowIndex, includeColors);
        }
        public DOMNode Document { get; set; }

        protected override void Paint(Graphics g, Rectangle clipBounds, Rectangle cellBounds,
            int rowIndex, DataGridViewElementStates cellState, object value, object formattedValue,
            string errorText, DataGridViewCellStyle cellStyle, DataGridViewAdvancedBorderStyle advancedBorderStyle,
            DataGridViewPaintParts paintParts)
        {

            //背景色
            bool selected = (cellState & DataGridViewElementStates.Selected) == DataGridViewElementStates.Selected;
            Color clr_background = !selected
                                       ? cellStyle.BackColor
                                       : cellStyle.SelectionBackColor;
            using (Brush bru = new SolidBrush(clr_background))
            {
                g.FillRectangle(bru, cellBounds);
            }

            //边框
            if ((paintParts & DataGridViewPaintParts.Border) != 0)
                PaintBorder(g, clipBounds, cellBounds, cellStyle, advancedBorderStyle);

            Rectangle rect_border = BorderWidths(advancedBorderStyle);
            Rectangle rect = cellBounds;
            rect.Offset(rect_border.X, rect_border.Y);
            rect.Width -= rect_border.Right;
            rect.Height -= rect_border.Bottom;

            if (value != null)
            {
                string text = value.ToString();
                if (text.Length > 0)
                    PaintText(g, rect, text, selected, cellStyle);
            }
            //自动换行

        }

        void PaintText(Graphics g, Rectangle rect, string text, bool selected, DataGridViewCellStyle cell_style)
        {
            Document = parser.parse_text(text);
            Point pt = rect.Location;
            foreach (var node in Document.Nodes)
            {
                Font font = get_node_font(node, cell_style);
                Color clr = get_node_color(node, selected, cell_style);
                Size size = TextRenderer.MeasureText(node.InnerText, font);
                pt.Y = rect.Top + (rect.Height - size.Height) / 2;
                if (clr.ToArgb() == Color.White.ToArgb() ||
                    clr.ToArgb() == Color.FromArgb(255, 0, 0, 0).ToArgb())
                {
                    using (Brush bru = new SolidBrush(clr))
                    {
                        g.DrawString(node.InnerText, font, bru, pt);
                    }
                }
                else
                {
                    using (Brush bru = new SolidBrush(clr))
                    {
                        PointF ptF = new PointF(pt.X, pt.Y);
                        SizeF sizeF = g.MeasureString(node.InnerText, font);
                        g.FillRectangle(bru, new RectangleF(ptF, sizeF));
                        g.DrawString(node.InnerText, font, Brushes.Black, pt);
                    }
                }

                pt.X += size.Width;
            }
        }

        Font get_node_font(DOMNode node, DataGridViewCellStyle cell_style)
        {
            Font font = cell_style.Font;
            if (node.Name == "font" && node.Attributes.ContainsKey("name") && node.Attributes.ContainsKey("size"))
            {
                float font_size = cell_style.Font.Size;
                float.TryParse(node.Attributes["size"], out font_size);
                font = new Font(node.Attributes["name"], font_size);
            }
            return font;
        }

        Color get_node_color(DOMNode node, bool selected, DataGridViewCellStyle cell_style)
        {
            Color clr = selected ? cell_style.SelectionForeColor : cell_style.ForeColor;
            if (node.Name == "body" && node.Attributes.ContainsKey("bgcolor"))
            {
                clr = ColorTranslator.FromHtml(node.Attributes["bgcolor"]);
                //if (selected)
                //    clr = Color.FromArgb(~(clr.ToArgb() & 0x00FFFFFF));
            }
            return clr;
        }
    }
    public class DataGridViewRTBEditingControl : RichTextBox, IDataGridViewEditingControl
    {

        public DataGridView EditingControlDataGridView
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public object EditingControlFormattedValue
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public int EditingControlRowIndex
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public bool EditingControlValueChanged
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public Cursor EditingPanelCursor
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public bool RepositionEditingControlOnValueChange
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public void ApplyCellStyleToEditingControl(DataGridViewCellStyle dataGridViewCellStyle)
        {
            throw new NotImplementedException();
        }

        public bool EditingControlWantsInputKey(Keys keyData, bool dataGridViewWantsInputKey)
        {
            throw new NotImplementedException();
        }

        public object GetEditingControlFormattedValue(DataGridViewDataErrorContexts context)
        {
            throw new NotImplementedException();
        }

        public void PrepareEditingControlForEdit(bool selectAll)
        {
            throw new NotImplementedException();
        }
    }

    public class DataGridViewRTBColumn : DataGridViewTextBoxColumn
    {
        public DataGridViewRTBColumn() : base()
        {
            CellTemplate = new DataGridViewRTBCell();
        }
    }
    public class DataGridViewRTBCell : DataGridViewTextBoxCell
    {
        public DataGridViewRTBCell() : base() { }
        public override void InitializeEditingControl(int rowIndex, object initialFormattedValue, DataGridViewCellStyle dataGridViewCellStyle)
        {
            base.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle);
            DataGridViewRTBEditingControl ctl = DataGridView.EditingControl as DataGridViewRTBEditingControl;
            ctl.WordWrap = true;
        }
        protected override void Paint(Graphics graphics, Rectangle clipBounds, Rectangle cellBounds, int rowIndex, DataGridViewElementStates cellState, object value, object formattedValue, string errorText, DataGridViewCellStyle cellStyle, DataGridViewAdvancedBorderStyle advancedBorderStyle, DataGridViewPaintParts paintParts)
        {
            //base.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, paintParts);
            ////背景色
            bool selected = (cellState & DataGridViewElementStates.Selected) == DataGridViewElementStates.Selected;
            Color clr_background = !selected
                                       ? cellStyle.BackColor
                                       : cellStyle.SelectionBackColor;
            using (Brush bru = new SolidBrush(clr_background))
            {
                graphics.FillRectangle(bru, cellBounds);
            }

            //边框
            if ((paintParts & DataGridViewPaintParts.Border) != 0)
                PaintBorder(graphics, clipBounds, cellBounds, cellStyle, advancedBorderStyle);

            Rectangle rect_border = BorderWidths(advancedBorderStyle);
            Rectangle rect = cellBounds;
            rect.Offset(rect_border.X, rect_border.Y);
            rect.Width -= rect_border.Right;
            rect.Height -= rect_border.Bottom;

            if (value != null)
            {
                string text = value.ToString();
                if (text.Length > 0)
                    PaintText(graphics, rect, text, selected, cellStyle);
            }

        }
        static DOMParser parser = new DOMParser();
        public DOMNode Document { get; set; }
        void PaintText(Graphics g, Rectangle rect, string text, bool selected, DataGridViewCellStyle cell_style)
        {
            Document = parser.parse_text(text);
            Point pt = rect.Location;
            foreach (var node in Document.Nodes)
            {
                Font font = get_node_font(node, cell_style);
                Color clr = get_node_color(node, selected, cell_style);
                Size size = TextRenderer.MeasureText(node.InnerText, font);
                pt.Y = rect.Top + (rect.Height - size.Height) / 2;

                if (clr.ToArgb() == Color.White.ToArgb() ||
                    clr.ToArgb() == Color.FromArgb(255, 0, 0, 0).ToArgb())
                {
                    using (Brush bru = new SolidBrush(clr))
                    {
                        g.DrawString(node.InnerText, font, bru, pt);
                    }
                }
                else
                {
                    using (Brush bru = new SolidBrush(clr))
                    {
                        PointF ptF = new PointF(pt.X, pt.Y);
                        SizeF sizeF = g.MeasureString(node.InnerText, font);
                        g.FillRectangle(bru, new RectangleF(ptF, sizeF));
                        g.DrawString(node.InnerText, font, Brushes.Black, pt);
                    }
                }

                pt.X += size.Width;
            }
        }

        Font get_node_font(DOMNode node, DataGridViewCellStyle cell_style)
        {
            Font font = cell_style.Font;
            if (node.Name == "font" && node.Attributes.ContainsKey("name") && node.Attributes.ContainsKey("size"))
            {
                float font_size = cell_style.Font.Size;
                float.TryParse(node.Attributes["size"], out font_size);
                font = new Font(node.Attributes["name"], font_size);
            }
            return font;
        }

        Color get_node_color(DOMNode node, bool selected, DataGridViewCellStyle cell_style)
        {
            Color clr = selected ? cell_style.SelectionForeColor : cell_style.ForeColor;
            if (node.Name == "body" && node.Attributes.ContainsKey("bgcolor"))
            {
                clr = ColorTranslator.FromHtml(node.Attributes["bgcolor"]);
                //if (selected)
                //    clr = Color.FromArgb(~(clr.ToArgb() & 0x00FFFFFF));
            }
            return clr;
        }

    }
}
