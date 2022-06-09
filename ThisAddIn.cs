using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Net;
using System.IO;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace OutlookAutoSig
{
    public partial class ThisAddIn
    {
        public Outlook.Application OutlookApp;
        private Outlook.Inspectors _inspectors;
        private Outlook.MailItem _mailItem;
        public Outlook.Explorer currentExplorer = null;
        public Dictionary<string, string> accountDictionary;
        public string currentSightml;
        public string username;
        public string sign;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //MessageBox.Show("更新测试版本13");
            bool nettest = WebRequestTest();
            if (nettest) 
            {
                OutlookApp = Globals.ThisAddIn.Application;
                accountDictionary = CreateSigByLongsysUser();
                OutlookApp.Startup += Application_Startup;
            }
            
        }
        private void Application_Startup()
        {
            _inspectors = OutlookApp.Inspectors;
            _inspectors.NewInspector += inspectors_NewInspector;
            currentExplorer = OutlookApp.ActiveExplorer();
            currentExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(CurrentExplorer_Event);
        }
        void inspectors_NewInspector(Outlook.Inspector inspector)
        {
            try
            {
                _mailItem = inspector.CurrentItem;
                if (!string.IsNullOrEmpty(_mailItem.Subject))
                {
                    //MessageBox.Show(_mailItem.Subject);
                    return;
                }
                Outlook.Account current_account = _mailItem.SendUsingAccount;
                string current_smtpAddress = current_account.SmtpAddress;
                if (current_smtpAddress.Contains("longsys.com") || current_smtpAddress.Contains("lexar.com"))
                {
                    //string sightml = accountDictionary[current_smtpAddress];
                    //MessageBox.Show(sightml);
                    //MessageBox.Show(sightml.Length.ToString());
                    //MessageBox.Show(sightml.Equals(@"null\n").ToString());
                    //_mailItem.HTMLBody += @"<br><hr />";
                    //_mailItem.HTMLBody += currentSightml;
                    if (currentSightml.Length > 6)
                    {
                        _mailItem.HTMLBody += @"<br><hr />";
                        _mailItem.HTMLBody += currentSightml;
                        //_mailItem.HTMLBody += @"<p>自动签名插件测试</p>";
                        //_mailItem.HTMLBody += @"<p>我当前的用户名是:"+myname+"</p>";
                    }
                    else
                    {
                        return;
                    }

                }
                else
                {
                    return;
                }
            }
            catch
            {
                return;
            }
        }
        private void CurrentExplorer_Event()
        {
            if (this.Application.ActiveExplorer().Selection.Count == 1
            && this.Application.ActiveExplorer().Selection[1] is Outlook.MailItem)
            {
                if (_mailItem != null)
                {
                    ((Outlook.ItemEvents_10_Event)_mailItem).Reply -= new Outlook.ItemEvents_10_ReplyEventHandler(MailItem_Reply);
                    ((Outlook.ItemEvents_10_Event)_mailItem).ReplyAll -= new Outlook.ItemEvents_10_ReplyAllEventHandler(MailItem_Reply);
                    ((Outlook.ItemEvents_10_Event)_mailItem).Forward -= new Outlook.ItemEvents_10_ForwardEventHandler(MailItem_Reply);
                }

                _mailItem = this.Application.ActiveExplorer().Selection[1];
                ((Outlook.ItemEvents_10_Event)_mailItem).Reply += new Outlook.ItemEvents_10_ReplyEventHandler(MailItem_Reply);
                ((Outlook.ItemEvents_10_Event)_mailItem).ReplyAll += new Outlook.ItemEvents_10_ReplyAllEventHandler(MailItem_Reply);
                ((Outlook.ItemEvents_10_Event)_mailItem).Forward += new Outlook.ItemEvents_10_ForwardEventHandler(MailItem_Reply);
            }
        }
        void MailItem_Reply(Object response, ref bool cancel)
        {

            Outlook.MailItem mailItem = (Outlook.MailItem)response;
            Outlook.Account current_account = mailItem.SendUsingAccount;
            string current_smtpAddress = current_account.SmtpAddress;
            if (current_smtpAddress.Contains("longsys.com") || current_smtpAddress.Contains("lexar.com"))
            {
                //string sightml = accountDictionary[current_smtpAddress];
                if (currentSightml.Length > 6)
                {
                    string oldHTMLBody = null;
                    oldHTMLBody = mailItem.HTMLBody;
                    /*Form1 form1 = new Form1();
                    form1.Test(mailItem.HTMLBody);
                    form1.Show();
                    form1.Text = "原始邮件内容";*/
                    mailItem.GetInspector.Activate();
                    /*Form1 form2 = new Form1();
                    form2.Test(mailItem.HTMLBody);
                    form2.Show();
                    form2.Text = "回复时邮件内容";*/
                    /*        string u = mailItem.Subject;
                            string x = mailItem.Sender.Name;
                            string r = mailItem.SenderName;*/
                    mailItem.HTMLBody = @"<br><hr />" + currentSightml + oldHTMLBody;
                }
                else
                {
                    return;
                }

            }
            else
            {
                return;
            }
        }
        public Dictionary<string, string> CreateSigByLongsysUser()
        {
            Dictionary<string, string> myDictionary = new Dictionary<string, string>();
            Dictionary<string, string> myDictionary1 = new Dictionary<string, string>();
            Outlook.Accounts accounts = OutlookApp.Session.Accounts;
            foreach (Outlook.Account account in accounts)
            {
                string smtpAddress = account.SmtpAddress;
                if (smtpAddress.Contains("longsys.com") || smtpAddress.Contains("lexar.com"))
                {
                    string[] strArray = smtpAddress.Split('@');
                    string username = strArray[0];
                    string sightm0 = GetSigByUsername(username,"0");
                    string sightm_name = username + "签名0";
                    if (sightm0.Length > 6) { 
                        myDictionary.Add(sightm_name, sightm0);
                    };
                    string sightm1 = GetSigByUsername(username, "1");
                    string sightm1_name = username + "签名1";
                    if (sightm1.Length>6) { 
                        myDictionary.Add(sightm1_name, sightm1);
                    };
                }
            }
            return myDictionary;
            //MessageBox.Show(builder.ToString());
            //MessageBox.Show(myDictionary.ToString());
        }
        public string GetSigByUsername(string username,string sign)
        {
            string key = "TKD3EPVg4DWsJ6L7";
            string iv = "eAFmBHVilJgOTqlO";
            string enmyname = AESEncrypt.EnCode(username, key, iv);
            string url = "http://ex.longsys.com:8001/querySiga";
            //string url = "http://localhost:8080/querySiga";
            Dictionary<string, string> data = new Dictionary<string, string>();
            data.Add("username", enmyname);
            data.Add("sign", sign);
            string sightml = Post(url, data);
            return sightml;
        }

        public static bool WebRequestTest()
        {
            string url = "http://ex.longsys.com:8001";
            try
            {
                WebRequest myRequest = WebRequest.Create(url);
                WebResponse myResponse = myRequest.GetResponse();
            }
            catch (WebException)
            {
                return false;
            }
            return true;
        }
        public string Post(string url, Dictionary<string, string> dic)
        {
            string result = "";
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "POST";
            req.ContentType = "application/x-www-form-urlencoded";
            #region 添加Post 参数  
            StringBuilder builder = new StringBuilder();
            int i = 0;
            foreach (var item in dic)
            {
                if (i > 0)
                    builder.Append("&");
                builder.AppendFormat("{0}={1}", item.Key, item.Value);
                i++;
            }
            byte[] data = Encoding.UTF8.GetBytes(builder.ToString());
            req.ContentLength = data.Length;
            using (Stream reqStream = req.GetRequestStream())
            {
                reqStream.Write(data, 0, data.Length);
                reqStream.Close();
            }
            #endregion
            HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
            Stream stream = resp.GetResponseStream();
            //获取响应内容  
            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
            {
                result = reader.ReadToEnd();
            }
            return result;
        }
        public class AESEncrypt
        {
            /// <summary>
            /// 加密数据
            /// </summary>
            /// <param name="data">明文数据</param>
            /// <param name="key">密钥</param>
            /// <param name="iv">偏移</param>
            /// <returns></returns>
            public static string EnCode(string data, string key, string iv)
            {
                try
                {
                    byte[] keyArray = Encoding.UTF8.GetBytes(key);
                    byte[] ivArray = Encoding.UTF8.GetBytes(iv);
                    byte[] toEncryptArray = Encoding.UTF8.GetBytes(data);
                    RijndaelManaged rDel = new RijndaelManaged();
                    rDel.Key = keyArray;
                    rDel.IV = ivArray;
                    rDel.Mode = CipherMode.CBC;
                    rDel.Padding = PaddingMode.PKCS7;
                    ICryptoTransform cTransform = rDel.CreateEncryptor();
                    byte[] resultArray = cTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);
                    return Convert.ToBase64String(resultArray, 0, resultArray.Length);
                }
                catch (System.Exception e)
                {
                    return null;
                }
            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 备注: Outlook不会再触发这个事件。如果具有
            //    在 Outlook 关闭时必须运行，详请参阅 https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
