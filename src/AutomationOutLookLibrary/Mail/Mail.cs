
using System;
using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
namespace AutomationOutLookLibrary
{
    public partial class OutlookApp
    {
        public void NewEmail()
        {
            mailItem = outlookApp.GetType().InvokeMember("CreateItem", BindingFlags.InvokeMethod, null, outlookApp, new object[] { OlItemType.olMailItem });
        }

        public void SetMailDisplay()
        {
            mailItem.GetType().InvokeMember("Display", BindingFlags.InvokeMethod, null, mailItem, new object[] { true });
        }

        public void SetMailAddAttachments(string attachmentPath)
        {
            object attachments = mailItem.GetType().InvokeMember("Attachments", BindingFlags.GetProperty, null, mailItem, new object[] { });

            attachments.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, attachments, new object[] { attachmentPath });
        }

        public void SetMailSendTo(string sendToEmailAdress)
        {
            mailItem.GetType().InvokeMember("To", BindingFlags.SetProperty, null, mailItem, new object[] { sendToEmailAdress });
        }

        public void SetMailSubject(string emailSubject)
        {
            mailItem.GetType().InvokeMember("Subject", BindingFlags.SetProperty, null, mailItem, new object[] { emailSubject });
            EmailTitle = emailSubject;
        }

        public void SetMailBodyFormat(OlBodyFormat olBodyFormat)
        {
            mailItem.GetType().InvokeMember("BodyFormat", BindingFlags.SetProperty, null, mailItem, new object[] { olBodyFormat });
        }

        public void SetMailHTMLBody(string contentHTML)
        {
            mailItem.GetType().InvokeMember("HTMLBody", BindingFlags.SetProperty, null, mailItem, new object[] { contentHTML + CommonHelper.ReadSignature() });
        }
        /// <summary>
        /// This function aim to add vote button in mail,the options is string separate with semicolon
        /// the value like add Yes button an No button example:"Yes;No;"
        /// </summary>
        /// <param name="votingOptions"></param>
        public void SetMailVotingOptions(string votingOptions)
        {
            mailItem.GetType().InvokeMember("VotingOptions", BindingFlags.SetProperty, null, mailItem, new object[] { votingOptions });
        }

        public void SetCC(string ccEmailAdress)
        {

            dynamic mail = mailItem;
            mail.CC = ccEmailAdress;
            //mailItem.GetType().InvokeMember("CC", BindingFlags.InvokeMethod, null, mailItem, new object[] { true });
        }

        public void SetBCC(string bccEmailAdress)
        {
            dynamic mail = mailItem;
            mail.BCC = bccEmailAdress;
            //mailItem.GetType().InvokeMember("BCC", BindingFlags.InvokeMethod, null, mailItem, new object[] { true });
        }

        public void SendMail()
        {

            if (!IsAutoSend)
            {
                SendThread();
            }
        }
        public void ShowEmail()
        {
            //sendMailThread = new Thread(new ThreadStart(() =>
            //{
            //    SetMailDisplay();
            //}));
            //sendMailThread.IsBackground = false;
            //sendMailThread.Start();
            new Action(() => { SetMailDisplay(); }).BeginInvoke(null, null);

            string title = string.Format("{0} - Message (HTML) ", EmailTitle);
            AutomationElement newEmailElement = Common.WaitForElementByName(100, title);
            while (newEmailElement == null)
            {
                try
                {
                    newEmailElement = Common.WaitForElementByName(100, title);
                }
                catch { }
            }
        }

        private void SendThread()
        {
            //sendMailThread = new Thread(new ThreadStart(() =>
            //{
            //    SetMailDisplay();
            //}));
            //sendMailThread.IsBackground = false;
            //sendMailThread.Start();
            this.ShowEmail();
            while (true)
            {
                try
                {
                    string title = string.Format("{0} - Message (HTML) ", EmailTitle);
                    var newEmailElement = Common.WaitForElementByName(100, title);

                    var btnSendCondition = new PropertyCondition(AutomationElement.NameProperty, "Send");
                    var btnInternalCondition = new PropertyCondition(AutomationElement.NameProperty, "Internal");

                    var sendElement = Common.WaitForElement(newEmailElement, btnSendCondition, 100);
                    var internalElement = Common.WaitForElement(newEmailElement, btnInternalCondition, 100);

                    XPoint p;
                    Common.GetCursorPos(out p);
                    Common.SetWindowPos(newEmailElement.Current.NativeWindowHandle, WindowsLayer.TOPMOST);

                    //internal click
                    var point = internalElement.GetClickablePoint();
                    Common.SetCursorPos((int)point.X, (int)point.Y);
                    Common.mouse_event(MouseArg.MOUSEEVENTF_LEFTDOWN | MouseArg.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

                    point = sendElement.GetClickablePoint();
                    Common.SetCursorPos((int)point.X, (int)point.Y);
                    Common.mouse_event(MouseArg.MOUSEEVENTF_LEFTDOWN | MouseArg.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

                    Common.SetCursorPos((int)p.X, (int)p.Y);

                    Thread.Sleep(1000);

                    int newEmailHandleId = Common.FindWindow(null, title);
                    if (newEmailHandleId == 0)
                    {
                        //sendMailThread.Abort();
                        //sendMailThread.DisableComObjectEagerCleanup();
                        //sendMailThread = null;
                        break;
                    }
                    else
                    {
                        int dialogHandle = Common.FindWindow(null, "Classify");
                        Common.SendMessage(dialogHandle, SendMsgArg.WM_ClOSE, 0, 0);
                        Common.SendMessage(dialogHandle, SendMsgArg.WM_DISTROY, 0, 0);
                        Common.SendMessage(dialogHandle, SendMsgArg.WM_QUIT, 0, 0);
                        //break;
                    }
                }
                catch
                {
                    Thread.Sleep(500);
                }
            }

        }


        //public bool SetOutlookHomePageHidden()
        //{
        //    //bool result = false;
        //    //Email.OutlookHomeScreen homeScreen = new Email.OutlookHomeScreen();
        //    //while (true)
        //    //{
        //    //    if (homeScreen.WaitForCreate(20))
        //    //    {
        //    //        homeScreen.btnMinimize.ClickSecureButton();
        //    //        result = true;
        //    //        break;
        //    //    }

        //    //    Thread.Sleep(500);
        //    //}

        //    //return result;

        //    return true;
        //}


    }
}
