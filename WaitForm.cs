using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Drawing.Text;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Threading;
using System.ComponentModel;
using com.sun.org.apache.bcel.@internal.classfile;

namespace PluginSalas
{
    public partial class WaitForm : Form
    {

        int progress = 0;

        [DllImport("gdi32.dll")]
        private static extern IntPtr AddFontMemResourceEx(IntPtr pbfont, uint cbfont, IntPtr pdv, [In] ref uint pcFonts);


        FontFamily ff;
        Font font;
        public WaitForm()
        {
            InitializeComponent();
        }

        private void loadFont()
        {
            byte[] fontArray = PluginSalas.Properties.Resources.HelvNeue75_W1G;
            int dataLength = PluginSalas.Properties.Resources.HelvNeue75_W1G.Length;

            IntPtr ptrData = Marshal.AllocCoTaskMem(dataLength);

            Marshal.Copy(fontArray, 0, ptrData, dataLength);

            uint cFonts = 0;

            AddFontMemResourceEx(ptrData, (uint)fontArray.Length, IntPtr.Zero, ref cFonts);

            PrivateFontCollection pfc = new PrivateFontCollection();

            pfc.AddMemoryFont(ptrData, dataLength);

            Marshal.FreeCoTaskMem(ptrData);

            ff = pfc.Families[0];
            font = new Font(ff, 15f, FontStyle.Regular);
        }

        private void AllocFont(Font f, Control c, float size)
        {
            FontStyle fontStyle = FontStyle.Regular;

            c.Font = new Font(ff, size, fontStyle);
        }


        private void PictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void WaitForm_Load(object sender, EventArgs e)
        {
            Code Snippet;
            CheckForIllegalCrossThreadCalls = false;
            loadFont();
            AllocFont(font, this.label1, 11);

        }

        public void ShowWait()
        {
            this.Show();
            Application.DoEvents();
        }

        public void CloseWait()
        {
            this.Close();
        }
     
    }
}
