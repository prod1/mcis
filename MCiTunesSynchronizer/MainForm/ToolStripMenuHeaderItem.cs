using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace MCiTunesSynchronizer
{
    class ToolStripMenuHeaderItem : ToolStripMenuItem
    {
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(this.BackColor);
            e.Graphics.DrawImage(this.Image, new PointF(5, 2));
            e.Graphics.DrawString(this.Text, new Font(this.Font, FontStyle.Bold | FontStyle.Italic), new SolidBrush(this.ForeColor), new PointF(24, 2));
        }
    }
}
