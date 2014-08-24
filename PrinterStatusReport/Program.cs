using System;
using System.Collections.Generic;
using System.Text;

using System.Net;
using System.IO;
using System.Net.Mail;
using System.Text.RegularExpressions;

namespace PrinterStatusReport
{
    class Program
    {
        static void Main(string[] args)
        {
            PrinterStatus PS = new PrinterStatus();

            //System.Threading.Thread.Sleep(100 * 1000);
        }
    }

    class MailAlert{
        private string smtp_user="alert.mail.sender@gmail.com";
        private string smtp_pwd = "alert.mail.pwd";

        private string mail_body = "";
        private string rec_default = "your.mail@gmail.com";
        private string rec_list = System.Environment.CurrentDirectory + @"\printer_alert_address.txt";
        private MailMessage msg;

        public MailAlert() {
            msg = new MailMessage(smtp_user,this.ReadList());
            msg.IsBodyHtml = false;
            msg.BodyEncoding = System.Text.Encoding.UTF8;
            msg.SubjectEncoding = System.Text.Encoding.UTF8;
            msg.Subject = "["+DateTime.Now.ToString("yyyy-MM-dd")+"] Printer Status";
            msg.Priority = MailPriority.Normal;
        }

        public void Send() {
            msg.Body = mail_body;

            SmtpClient SC = new SmtpClient("smtp.gmail.com",587);
            SC.Credentials = new NetworkCredential(smtp_user, smtp_pwd);
            SC.EnableSsl = true;
            SC.Send(this.msg);
            SC.Dispose();
        }

        public void SetMailBody(string msg){
            this.mail_body = msg; 
        }
        
        public string ReadList(){
            if (File.Exists(rec_list))
            {
                string line="";
                List<string> mail_addr=new List<string>();
                StreamReader file = new StreamReader(rec_list);
                while ((line = file.ReadLine())!=null) {
                    if (line.IndexOf("#") >= 0) { continue; }
                    if (line.Length <= 4) { continue; }//a@a.c
                    mail_addr.Add(line.Trim());
                }
                file.Close();
                return String.Join(",",mail_addr.ToArray());
            }
            else {
                return rec_default;
            }
        }
    }

    class PrinterStatus {
        //Epson AcuLaser CX17NF v01.00.28
        private string URL_Toner=@"http://[your printer ip]/status/statsupl.asp";
        private string URL_Count = @"http://[your printer ip]/setting/counter.asp";


        public PrinterStatus() {
            string mail_body = "Printer Status\r\n\r\n";
            try
            {
                string html_toner = this.GetHtml(URL_Toner); ;
                Dictionary<string, string> Toner = this.Parser(html_toner, "toner");

                foreach (KeyValuePair<string, string> KP in Toner)
                {
                    Console.WriteLine(KP.Key + " : " + KP.Value);
                    mail_body += KP.Key + " : " + KP.Value + "\r\n";
                }

                mail_body += "\r\n\r\n";
                string html_count = this.GetHtml(URL_Count);
                Dictionary<string, string> Count = this.Parser(html_count, "count");

                foreach (KeyValuePair<string, string> KP in Count)
                {
                    Console.WriteLine(KP.Key + " : " + KP.Value);
                    mail_body += KP.Key + " : " + KP.Value + "\r\n";
                }
            }
            catch (Exception)
            {
                mail_body = "Get Printer Status Error";
            }
            finally
            {

                MailAlert MA = new MailAlert();
                MA.SetMailBody(mail_body+"\r\n\r\nTime : "+DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                MA.Send();
            }
        }

        public string GetHtml(string url){
            string html = "";

            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Timeout = 5000;
            //if you need zh-tw language Data
            //req.Headers.Add("Accept-Language", "zh-TW,zh;q=0.8,en-US;q=0.6,en;q=0.4");
            HttpWebResponse response = (HttpWebResponse)req.GetResponse();
            if (response.StatusCode == HttpStatusCode.OK)
            {
                StreamReader sr = new StreamReader(response.GetResponseStream());
                html = sr.ReadToEnd();
            }
            response.Close();

            return html;
        }

        private Dictionary<string, string> Parser(string html,string page_type) {
            Dictionary<string, string> rt=new Dictionary<string,string>();

            if (page_type.Equals("toner"))//toner status
            {
                int tmp_s = html.IndexOf("Components");
                string[] data = html.Substring(tmp_s, html.IndexOf("<center><table>") - tmp_s).Split(new string[] { "</tr>" }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string row in data)
                {
                    //Console.WriteLine(row);
                    MatchCollection MC = Regex.Matches(row.Replace("\n", "").Replace("\r", ""), "(?<item>[^>.]+)</font>.*size=-1>(.*>){0,1}(?<val>[^>.]+)</font>");
                    foreach (Match match in MC)
                    {
                        //Console.WriteLine(match.Groups["item"]+" "+match.Groups["val"]);
                        rt.Add(match.Groups["item"].Value, match.Groups["val"].Value.Replace("&#37;", " %"));
                    }
                }
            }
            else if (page_type.Equals("count")) {//print count
                int tmp_s = html.IndexOf("<!--Usage Counters-->");
                string[] data = html.Substring(tmp_s, html.IndexOf("</tr>-->") - tmp_s).Split(new string[] { "</tr>" }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string row in data) {
                    //Console.WriteLine(row);
                    MatchCollection MC = Regex.Matches(row, "size=-1>(?<item>[^<]+).*size=-1>(?<val>[]0-9]+)");
                    foreach (Match match in MC) {
                        //Console.WriteLine(match.Groups["item"] + " : " + match.Groups["val"]);
                        rt.Add(match.Groups["item"].Value, match.Groups["val"].Value);
                    }
                }
            }

            return rt;
        }
    }
}
