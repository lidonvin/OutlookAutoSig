using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static OutlookAutoSig.ThisAddIn;

namespace OutlookAutoSig
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string username = Globals.ThisAddIn.username;
            string sign = Globals.ThisAddIn.sign;
            string name = textBox1.Text;
            string ename = textBox2.Text;
            string position = textBox3.Text;
            string department = textBox4.Text;
            string mobile = textBox5.Text;
            string tel = textBox6.Text;
            if (name.Length==0 && ename.Length == 0 && position.Length == 0 && department.Length == 0 && mobile.Length == 0 && tel.Length == 0)
            {
                MessageBox.Show("错误，数据不能为空，您至少需要填写一项！");
            }
            else {
                UpdateSigByUsername(username, name, ename, position, department, mobile, tel, sign);
                MessageBox.Show("修改完成，请重启Outlook生效。");
                this.Close();
            }          
        }
        public void UpdateSigByUsername(string username, string name, string ename, string position, string department, string mobile, string tel, string sign)
        {
            string key = "TKD3EPVg4DWsJ6L7";
            string iv = "eAFmBHVilJgOTqlO";
            string enmyname = AESEncrypt.EnCode(username, key, iv);
            string url = "http://ex.longsys.com:8001/updateSiga";
            //string url = "http://localhost:8080/updateSiga";
            Dictionary<string, string> data = new Dictionary<string, string>();
            data.Add("username", enmyname);
            data.Add("name", name);
            data.Add("ename", ename);
            data.Add("position", position);
            data.Add("department", department);
            data.Add("mobile", mobile);
            data.Add("tel", tel);
            data.Add("sign", sign);
            string sightml = Post(url, data);
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
    }
}
