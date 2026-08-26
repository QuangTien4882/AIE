using System.Drawing;
using System.Windows.Forms;

namespace AIE.ExcelAddIn.Helpers;

public static class UIHelper
{
    public static void ApplyStyle(TabControl tabControl)
    {
        tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
        tabControl.DrawItem -= TabControl_DrawItem; // Ensure we don't attach multiple times
        tabControl.DrawItem += TabControl_DrawItem;
    }

    private static void TabControl_DrawItem(object? sender, DrawItemEventArgs e)
    {
        if (sender is not TabControl tabControl) return;

        var g = e.Graphics;
        var tabArea = tabControl.GetTabRect(e.Index);

        if (e.Index == tabControl.SelectedIndex)
        {
            using (var brush = new SolidBrush(Color.FromArgb(0, 120, 215)))
            {
                g.FillRectangle(brush, tabArea);
            }
            TextRenderer.DrawText(g, tabControl.TabPages[e.Index].Text, tabControl.Font, tabArea, Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
        else
        {
            g.FillRectangle(Brushes.White, tabArea);
            TextRenderer.DrawText(g, tabControl.TabPages[e.Index].Text, tabControl.Font, tabArea, Color.Black, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
    }

    public static void ApplyStyle(DataGridView dgv)
    {
        // 1. Enable DoubleBuffering
        typeof(DataGridView).InvokeMember(
            "DoubleBuffered",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty,
            null,
            dgv,
            new object[] { true });

        // 2. Set Header Style
        dgv.EnableHeadersVisualStyles = false;
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(235, 235, 235);
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
        dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Bold);
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
        dgv.ColumnHeadersDefaultCellStyle.Padding = new Padding(0, 5, 0, 5);
        dgv.ColumnHeadersHeight = 45;
        dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;

        // 3. Row Styles and Wrapping
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
        dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
        dgv.RowHeadersVisible = true;
        dgv.RowHeadersWidth = 25;
        dgv.AllowUserToResizeRows = true;
        dgv.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
        dgv.RowTemplate.Resizable = DataGridViewTriState.True;
        dgv.RowTemplate.Height = 32;
        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 248, 252);
    }
}
