using OpenPop.Mime;
using OpenPop.Pop3;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Telegram.Bot;
using Telegram.Bot.Types.Enums;

namespace ReadGmail
{
    internal class Program
    {
        private static readonly TelegramBotClient Bot = new TelegramBotClient("Bot token");
        private static Timer _timer = null;
        static string fileName = "C:\\Sample.txt";
        static void Main(string[] args)
        {
            // it runs every 3 minutes 
            _timer = new Timer(OnTimeEvent, null, 0, 180000);
            Console.ReadKey();
        }

        // Read last email count from txt
        static int ReadTxt()
        {
            string line = "";
            try
            {
                //Pass the file path and file name to the StreamReader constructor
                StreamReader sr = new StreamReader(fileName);
                //Read the first line of text
                //Continue to read until you reach end of file
                if (line != null)
                {
                    line = sr.ReadLine();
                }
                //close the file
                sr.Close();
                return Convert.ToInt32(line);
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
                return 0;
            }

        }

        // Write last email count to txt
        static void WriteTxt(string count)
        {
            try
            {
                //Pass the filepath and filename to the StreamWriter Constructor
                StreamWriter sw = new StreamWriter(fileName);
                //Write a line of text
                sw.Write(count.ToString());
                //Close the file
                sw.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }

        }

        private static void OnTimeEvent(Object o)
        {
          
            if (!File.Exists(fileName))
            {
                // Create the file and use streamWriter to write text to it.
                //If the file existence is not check, this will overwrite said file.
                //Use the using block so the file can close and vairable disposed correctly
                using (StreamWriter writer = File.CreateText(fileName))
                {
                    writer.Write("0");
                }
            }
            Pop3Client pop3Client = new Pop3Client();
            pop3Client.Connect("pop.gmail.com", 995, true);
            pop3Client.Authenticate("email", "şifre");
            int count = pop3Client.GetMessageCount();
            int previousCount = ReadTxt();
            int current = count - previousCount;
            WriteTxt(count.ToString());
            List<EmailModel> emails = new List<EmailModel>();
            for (int i = count; i > count - current; i--)
            {
                Message message = pop3Client.GetMessage(i);
                EmailModel email = new EmailModel()
                {
                    MessageNumber = i,
                    From = string.Format("<a href = 'mailto:{1}'>{0}</a>", message.Headers.From.DisplayName, message.Headers.From.Address),
                    Subject = message.Headers.Subject,
                    DateSent = message.Headers.DateSent
                };

                // Part of email subject you can replace word1,word2 with the keywords you need 
                if (!email.Subject.ToLower().Contains("word1") && !email.Subject.ToLower().Contains("word2"))
                {
                    MessagePart body = message.FindFirstHtmlVersion();
                    if (body != null)
                    {
                        email.Body = body.GetBodyAsText();
                    }
                    else
                    {
                        body = message.FindFirstPlainTextVersion();
                        if (body != null)
                        {
                            email.Body = body.GetBodyAsText();
                        }
                    }
                    email.BodyCopy = email.Body;

                    //Part of body subject you can replace w1, AW1, AW2 with the keywords you need

                    while (email.BodyCopy.Contains("w1"))
                    {

                        int index = email.BodyCopy.IndexOf("w1");
                        if (email.BodyCopy.Length < (index + 12))
                        {
                            email.GatewayId.Add("Not Found");
                        }
                        else
                        {
                            email.GatewayId.Add(email.BodyCopy.Substring(index + 12, 36).Replace("</td>\r\n<td align=\"left\"", ""));
                        }
                        email.BodyCopy = email.BodyCopy = email.BodyCopy.Substring(index + 12, email.BodyCopy.Length - (index + 12));
                        if (email.Body.Contains("AW1"))
                        {
                            email.Firma = "AW1";
                        }
                        else if (email.Body.Contains("AW2") || email.Body.Contains("AW22"))
                        {
                            email.Firma = "AW2";
                        }
                        else if (email.Body.Contains("AW3"))
                        {
                            email.Firma = "AW3";
                        }

                    }
                    emails.Add(email);


                }

            }

            //Get All Last Messages and Send to telegram bot

            foreach (var item in emails.OrderBy(x => x.DateSent))
            {
                item.GatewayList = "";
                foreach (var items in item.GatewayId)
                {
                    item.GatewayList = item.GatewayList + items + "\n";
                }
                if (!String.IsNullOrEmpty(item.GatewayList))
                {
                    string texto = @"<b>⚠️ " + item.Firma + " ⚠️</b>" +
                        Environment.NewLine +
                        Environment.NewLine +
                        @"<code>" + "Subject: " + item.Subject + "</code>" + Environment.NewLine +
                        @"<code>" + "W1: " + Environment.NewLine + item.GatewayList + "</code>" +
                        Environment.NewLine;
                    Console.WriteLine(texto);
                    Bot.SendTextMessageAsync("Group ID", texto, parseMode: ParseMode.Html);
                }
            }
        }
       
    }
}
