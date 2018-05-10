using System;
using System.Diagnostics;
using System.Threading;

namespace AutomationOutLookLibrary
{
    public partial class OutlookApp : IDisposable
    {
        internal object outlookApp;
        internal object mailItem;
        //internal Thread sendMailThread;
        private bool IsAutoSend = false;
        private string EmailTitle = string.Empty;

        public OutlookApp(bool isAutoSend = true)
        {
            IsAutoSend = isAutoSend;
            #region
            bool isStartOutlook = false;

            Process[] outlookList = Process.GetProcessesByName("Outlook");

            if (outlookList.Length == 0)
            {
                Process.Start("Outlook");
                isStartOutlook = true;
            }

            while (true)
            {
                if (isStartOutlook)
                {
                    Thread.Sleep(15000);
                }

                //if (SetOutlookHomePageHidden())
                //{
                //    break;
                //}
                break;
            }
            #endregion

            if (isAutoSend)
            {
                SendMail();
            }
            outlookApp = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application"));
        }
        public void Dispose()
        {
            //if (sendMailThread != null)
            //{
            //    sendMailThread.Abort();
            //    sendMailThread.DisableComObjectEagerCleanup();
            //    sendMailThread = null;
            //}
            outlookApp = null;
            mailItem = null;
        }
    }
}
