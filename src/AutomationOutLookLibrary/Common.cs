using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Automation;

namespace AutomationOutLookLibrary
{
   internal class Common
    {
        public static AutomationElement WaitForElement(AutomationElement parent, Condition condition, int millsecondTimeout)
        {

            var waitTime = 0;
            AutomationElement element = null;
            try
            {
                element = parent.FindFirst(TreeScope.Children, condition);
            }
            catch (Exception)
            {
                throw;
            }
            while (element == null)
            {
                if (waitTime >= millsecondTimeout) break;
                System.Threading.Thread.Sleep(500);
                waitTime += 500;
                element = parent.FindFirst(TreeScope.Subtree, condition);
            }

            return element;

        }
        public static AutomationElement WaitForElementByName(int millisecondTimeout, string nameTexts)
        {
            var parent = AutomationElement.RootElement;
            var condition = new PropertyCondition(AutomationElement.NameProperty, nameTexts);
            return WaitForElement(parent, condition, millisecondTimeout);
        }

        [DllImport("user32.dll")]
        public static extern int SetWindowPos(int hwnd, WindowsLayer hWndInsertAfter, int x = 0, int y = 0, int cx = 0, int cy = 0, int wFlag = 3);


        [DllImport("user32.dll", EntryPoint = "FindWindowA")]
        public static extern int FindWindow(string lpClassName, string lpWidnowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int FindWindowEx(int parent /*HWND*/, int next /*HWND*/, string sClassName, int sWindowTitle);

        [DllImport("user32.dll", EntryPoint = "SendMessageA")]
        public static extern int SendMessage(int hwnd, SendMsgArg msg, int wparam, int lparam);

        [DllImport("user32.dll")]
        public static extern int SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        public static extern void mouse_event(MouseArg dwflag, int x, int y, int cbuttons, int dwextrainfo);


        [DllImport("user32.dll")]
        public static extern int GetCursorPos(out XPoint p);


        [DllImport("user32.dll")]
        private static extern int ShowWindow(int hwnd, CmdShow nCmdShow);

        public static int ShowWindowEx(int hwnd, CmdShow nCmdShow)
        {

            ShowWindow(hwnd, nCmdShow);
            return ShowWindow(hwnd, nCmdShow);
        }
    }

    internal enum CmdShow
    {
        SW_HIDE = 0,
        SW_MAXIMIZE = 3,
        SW_MINIMIZE = 6,
        SW_RESTORE = 9,
        SW_SHOW = 5,
        SW_SHOWDEFAULT = 10,
        SW_SHOWMAXIMIZED = 2,
        SW_SHOWMINNOACTIVE = 7,
        SW_SHOWNA = 8,
        SW_SHOWNOACTIVE = 4,
        SW_SHOWNORMAL = 1
    }


    /// <summary>
    /// Define Mouse Behavior Type
    /// </summary>
    internal enum MouseArg
    {
        MOUSEEVENTF_MOVE = 0X0001,
        MOUSEEVENTF_LEFTDOWN = 0X0002,
        MOUSEEVENTF_LEFTUP = 0X0004,
        MOUSEEVENTF_RIGHTDOWN = 0X0008,
        MOUSEEVENTF_RIGHTUP = 0X0010,
        MOUSEEVENTF_MIDDLEDOWN = 0X0020,
        MOUSEEVENTF_MIDDLEUP = 0X0040,
        MOUSEEVENTF_ABSOLUTE = 0X8000
    }


    internal enum WindowsLayer
    {
        TOPMOST = -1,
        TOP = 0,
        NOTOPMOST = -2,
        DOWN = 1
    }

    /// <summary>
    /// message type
    /// </summary>
    internal enum SendMsgArg
    {
        #region keyboard
        WM_KEYDOWN = 0X100,
        WM_KEYUP = 0X101,
        WM_CHAR = 0X0102,
        WM_SYSKEYDOWN = 0X0104,
        WM_SYSKEYUP = 0X0105,
        WM_SYSCHAR = 0X0106,
        #endregion

        #region mouse
        BM_CLICK = 0XF5,
        WM_LBUTTONDOWN = 0X201,
        WM_LBUTTONUP = 0X202,

        WM_RBUTTONDOWN = 0X204,
        WM_RBUTTONUP = 0X205,

        WM_LBUTTONDBCLICK = 0X203,
        WM_RBUTTONDBCLICK = 0X206,

        WM_MOUSEWHEEL = 0X020A,
        #endregion

        #region text

        WM_SETTEXT = 0X0C,
        WM_GETTEXT = 0X0D,
        WM_CUT = 0X0300,
        WM_COPY = 0X0301,
        WM_PASTE = 0X0302,
        WM_CLEAR = 0X0303,
        WM_UNDO = 0X0304,

        #endregion

        #region windows

        WM_ClOSE = 0X0010,
        WM_DISTROY = 0X0002,
        WM_QUIT = 0X0012,

        #endregion


        #region command
        WM_SYSCOMMAND = 0X0112,
        WM_COMMAND = 0X0111
        //SC_MAXIMIZE SC_MINIMIZE SC_CLOSE
        #endregion
    }

    public struct XPoint
    {
        public int X;
        public int Y;
    }

}
