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

            Process p = Process.Start("C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE", " /S C:\\Users\\User\\Desktop\\123.ppsx");
            //Process p = Process.Start("C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE", " /S "+args[0]);

            p.WaitForInputIdle();

            int left = 100;
            int top = 100;
            int width = 384;
            int height = 216;

            while (!p.HasExited)
            {
                p.Refresh();
                IntPtr hwnd = p.MainWindowHandle;
                if (hwnd.ToInt32() != 0)
                {
                    // BOOL MoveWindow(HWND hWnd, int x, int y, int nWidth, int nHeight, BOOL bRepaint = TRUE);//

                    MoveWindow(hwnd , left, top, width, height,true);
                    //SetWindowPos(hwnd, HWND_TOPMOST, left, top, width, height, 0);
                    break;
                }
            }
        }
    }
}
