using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace WindowPPT
{
    class Program
    {
        // HWND Constants
        static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);

        // P/Invoke
        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X,
           int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int MoveWindow(IntPtr hWnd, int x, int y, int nWidth, int nHeight, bool BRePaint);

        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int LPAR);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();
        //Sets window attributes
        [DllImport("USER32.DLL")]
        public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        //Gets window attributes
        [DllImport("USER32.DLL")]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("USER32.DLL")]

        public static extern bool SetWindowRgn(IntPtr hWnd, IntPtr hRgn, bool redraw);

        [DllImport("USER32.DLL", CharSet = CharSet.Auto)]
        private static extern int ShowWindow(System.IntPtr hWnd,int nCmdShow);

        [DllImport("gdi32.dll")]

        static extern IntPtr CreateRoundRectRgn(int x1, int y1, int x2, int y2, int cx, int cy);

        public static int GWL_STYLE = -16;
        public static int WS_CHILD = 0x40000000; //child window
        public static int WS_BORDER = 0x00800000; //window with border
        public static int WS_DLGFRAME = 0x00400000; //window with double border but no title
        public static int WS_CAPTION = WS_BORDER | WS_DLGFRAME; //window with a title bar

        public const int SW_SHOWNORMAL = 1;
        public const int SW_RESTORE = 9;
        public const int SW_SHOW = 5;

        static void Main(string[] args)
        {

            //if (args != null)
            //{
            //    int argsLength = args.Length;
            //    Console.WriteLine(argsLength);
            //    for (int i = 0; i < argsLength; i++)
            //    {
            //        Console.WriteLine(args[i]);
            //    }
            //}

            Process p = Process.Start("C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE", " /S C:\\Users\\User\\Desktop\\123.pptx");
            //Process p = Process.Start("C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE", " /S "+args[0]);

            p.WaitForInputIdle();

            int left = 100;
            int top = 100;
            int width = 320;
            int height = 180;

            while (!p.HasExited)
            {
                p.Refresh();
                IntPtr hwnd = p.MainWindowHandle;
                if (hwnd.ToInt32() != 0)
                {
                    SetWindowRgn(hwnd, CreateRoundRectRgn(0, 30, 1280, 800, 0, 0), true);

                    //ShowWindow(hwnd, SW_SHOW);
                    //MoveWindow(hwnd , left, top, width, height,true);
                    //SetWindowPos(hwnd, HWND_TOPMOST, left, top, width, height, 0);
                    break;
                }
            }
        }
    }
}
