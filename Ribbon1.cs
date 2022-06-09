using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAutoSig
{
    public partial class Ribbon1
    {
        public Outlook.Application OutlookApp;
        public Dictionary<string, string> accountDictionary;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            OutlookApp = Globals.ThisAddIn.OutlookApp;
            accountDictionary = Globals.ThisAddIn.accountDictionary;
            if (accountDictionary != null) 
            {
                foreach (string k in accountDictionary.Keys)
                {
                    RibbonDropDownItem item = Factory.CreateRibbonDropDownItem();
                    item.Label = k;
                    comboBox1.Items.Add(item);
                }
                comboBox1.Text = accountDictionary.FirstOrDefault().Key;
                string signame = comboBox1.Text;
                string sign = signame.Substring(signame.Length - 1);
                string[] strArray = signame.Split('签');
                string username = strArray[0];
                Globals.ThisAddIn.currentSightml = accountDictionary[signame];
                Globals.ThisAddIn.username = username;
                Globals.ThisAddIn.sign = sign;
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            string email = comboBox1.Text;
            Form1 fm = new Form1();
            string sightml = (accountDictionary[email].Length > 6) ? accountDictionary[email] : "该邮箱未查询到用户相关信息，将使用本地签名。";
            fm.ShowHtml(sightml);
            fm.Show();
            fm.Text = email;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            string signame = comboBox1.Text;
            Form2 form2 = new Form2();
            form2.Show();
            form2.Text = "修改签名-"+ signame;
        }

        private void comboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            string signame = comboBox1.Text;
            string sign = signame.Substring(signame.Length - 1);
            string[] strArray = signame.Split('签');
            string username = strArray[0];
            Globals.ThisAddIn.currentSightml = accountDictionary[signame];
            Globals.ThisAddIn.username = username;
            Globals.ThisAddIn.sign = sign;
            //MessageBox.Show(username+"——"+ sign);
        }

    }
}
