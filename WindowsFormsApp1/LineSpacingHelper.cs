using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class RichTextBoxHelper
{
    private const int EM_SETPARAFORMAT = 0x447;
    private const int PFM_LINESPACING = 0x00000100;
    private const int SCF_SELECTION = 0x0001;

    [StructLayout(LayoutKind.Sequential)]
    private struct PARAFORMAT
    {
        public int cbSize;
        public uint dwMask;
        public short wNumbering;
        public short wReserved;
        public int dxStartIndent;
        public int dxRightIndent;
        public int dxOffset;
        public short wAlignment;
        public short cTabCount;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)]
        public int[] rgxTabs;
        public int dySpaceBefore;
        public int dySpaceAfter;
        public int dyLineSpacing;
        public short sStyle;
        public byte bLineSpacingRule;
        public byte bOutlineLevel;
        public short wShadingWeight;
        public short wShadingStyle;
        public short wNumberingStart;
        public short wNumberingStyle;
        public short wNumberingTab;
        public short wBorderSpace;
        public short wBorderWidth;
        public short wBorders;
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, ref PARAFORMAT lParam);

    public static void SetLineSpacing(RichTextBox rtb, float spacing)
    {
        if (spacing < 1.0f) spacing = 1.0f; // Prevents values below 1.0

        PARAFORMAT pf = new PARAFORMAT();
        pf.cbSize = Marshal.SizeOf(pf);
        pf.dwMask = PFM_LINESPACING;
        pf.bLineSpacingRule = 4; // 4 = Exact spacing (twips)
        pf.dyLineSpacing = (int)(spacing * 20); // Convert to twips (1 point = 20 twips)

        SendMessage(rtb.Handle, EM_SETPARAFORMAT, (IntPtr)SCF_SELECTION, ref pf);
    }
}
