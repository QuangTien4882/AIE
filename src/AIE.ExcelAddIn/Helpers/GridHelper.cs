using System;
using System.Drawing;
using System.Windows.Forms;

namespace AIE.ExcelAddIn.Helpers
{
    public static class GridHelper
    {
        public static void PaintMergedHeader(object sender, DataGridViewCellPaintingEventArgs e, DataGridView dgv, 
            int startCol1, int endCol1, string header1,
            int startCol2, int endCol2, string header2)
        {
            if (e.RowIndex == -1 && e.ColumnIndex >= 0)
            {
                e.Handled = true;

                // Determine if this column is part of a merged header
                bool isMerged1 = e.ColumnIndex >= startCol1 && e.ColumnIndex <= endCol1;
                bool isMerged2 = e.ColumnIndex >= startCol2 && e.ColumnIndex <= endCol2;

                // Save original clip and set clip to current cell to prevent bleeding
                Region oldClip = e.Graphics.Clip;
                e.Graphics.SetClip(e.CellBounds);

                // 1. Draw background
                using (Brush backBrush = new SolidBrush(e.CellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(backBrush, e.CellBounds);
                }

                // 2. Draw border
                using (Pen gridPen = new Pen(dgv.GridColor))
                {
                    // Bottom border (always drawn)
                    e.Graphics.DrawLine(gridPen, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right, e.CellBounds.Bottom - 1);
                    // Right border (always drawn)
                    e.Graphics.DrawLine(gridPen, e.CellBounds.Right - 1, e.CellBounds.Top, e.CellBounds.Right - 1, e.CellBounds.Bottom);
                    // Top border (always drawn)
                    e.Graphics.DrawLine(gridPen, e.CellBounds.Left, e.CellBounds.Top, e.CellBounds.Right, e.CellBounds.Top);
                    
                    if (isMerged1 || isMerged2)
                    {
                        // Draw horizontal line in the middle
                        int midY = e.CellBounds.Top + (e.CellBounds.Height / 2);
                        e.Graphics.DrawLine(gridPen, e.CellBounds.Left, midY, e.CellBounds.Right, midY);
                    }
                }

                // 3. Draw text
                using (Brush foreBrush = new SolidBrush(e.CellStyle.ForeColor))
                {
                    StringFormat format = new StringFormat
                    {
                        Alignment = StringAlignment.Center,
                        LineAlignment = StringAlignment.Center,
                        FormatFlags = StringFormatFlags.NoClip, // Allow wrapping
                        Trimming = StringTrimming.EllipsisCharacter
                    };

                    if (isMerged1 || isMerged2)
                    {
                        // Draw sub-header in the bottom half
                        Rectangle subRect = new Rectangle(e.CellBounds.Left, e.CellBounds.Top + e.CellBounds.Height / 2, e.CellBounds.Width, e.CellBounds.Height / 2);
                        
                        string text = e.Value?.ToString() ?? "";
                        if (text.StartsWith("ĐM ")) text = text.Substring(3);
                        if (text.StartsWith("CP ")) text = text.Substring(3);
                        if (text == "Khác (%)") text = "Khác";

                        e.Graphics.DrawString(text, e.CellStyle.Font, foreBrush, subRect, format);

                        // Draw main header across the group, clipping will naturally slice it
                        int startCol = isMerged1 ? startCol1 : startCol2;
                        int endCol = isMerged1 ? endCol1 : endCol2;
                        string mainHeader = isMerged1 ? header1 : header2;

                        // Calculate consistent mainRect for the entire merged group
                        Rectangle firstRect = dgv.GetCellDisplayRectangle(startCol, -1, false);
                        Rectangle lastRect = dgv.GetCellDisplayRectangle(endCol, -1, false);
                        
                        Rectangle mainRect = new Rectangle(
                            firstRect.Left, 
                            e.CellBounds.Top, 
                            lastRect.Right - firstRect.Left, 
                            e.CellBounds.Height / 2);
                        
                        e.Graphics.DrawString(mainHeader, e.CellStyle.Font, foreBrush, mainRect, format);
                    }
                    else
                    {
                        // Normal header, draw in the full bounds
                        e.Graphics.DrawString(e.Value?.ToString(), e.CellStyle.Font, foreBrush, e.CellBounds, format);
                    }
                }
                // Restore clip
                e.Graphics.Clip = oldClip;
                e.Handled = true;
            }
        }
    }
}
