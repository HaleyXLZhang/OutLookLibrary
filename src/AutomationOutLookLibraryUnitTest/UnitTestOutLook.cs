using AutomationOutLookLibrary;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;

namespace AutomationOutLookLibraryUnitTest
{
    [TestClass]
    public class UnitTestOutLook
    {
        [TestMethod]
        public void TestSendMail()
        {
            for (int i = 0; i < 1; i++)
            {
                string content = "附件为" + DateTime.Now.ToString("yyyyMMdd") + " 数据，请查阅，谢谢！";
                content = "各收件人，<br/>\r\n<br/>请重点关注以下内容：<br/>\r\n<br/>" + content + "<br/>\r\n<br/>  <br/>\r\n<br/>此邮件为系统自动邮件通知，请不要直接进行回复！谢谢。！";

                content = content + "<br/>\r\n<br/>";

                using (OutlookApp outlookApp = new OutlookApp(false))
                {
                    outlookApp.NewEmail();
                    outlookApp.SetMailSendTo("tanker.z.l.tan@noexternalmail.hsbc.com");
                    outlookApp.SetCC("tanker.z.l.tan@noexternalmail.hsbc.com");
                    outlookApp.SetMailSubject(DateTime.Now.ToString());
                    outlookApp.SetMailBodyFormat(OlBodyFormat.olFormatHTML);
                    outlookApp.SetMailHTMLBody(content);
                    //outlookApp.SetMailAddAttachments(@"C:\Publish\OAMOCR.Wpf.zip");
                    //outlookApp.SetMailAddAttachments(@"C:\Publish\OAMOCR.Wpf - Copy.zip");
                    //outlookApp.SetMailVotingOptions("Yes;No;");

                    outlookApp.ShowEmail();
                }
            }

        }
        [TestMethod]
        public void TestStartOutlook()
        {
            Process[] outlookList = Process.GetProcessesByName("Outlook");


        }


    }
}
