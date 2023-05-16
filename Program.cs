using System;
using System.Globalization;
using System.IO;
using EAGetMail; //add EAGetMail namespace

namespace receiveemail
{
    class Program
    {
        // Generate an unqiue email file name based on date time
        static string _generateFileName(int sequence)
        {
            DateTime currentDateTime = DateTime.Now;
            return string.Format("{0}-{1:000}-{2:000}.eml",
                currentDateTime.ToString("yyyyMMddHHmmss", new CultureInfo("en-US")),
                currentDateTime.Millisecond,
                sequence);
        }

        static void Main(string[] args)
        {
            try
            {
                // Create a folder named "inbox" under current directory
                // to save the email retrieved.
                string localInbox = string.Format("{0}\\inbox", Directory.GetCurrentDirectory());
                // If the folder is not existed, create it.
                if (!Directory.Exists(localInbox))
                {
                    Directory.CreateDirectory(localInbox);
                }

                string user = Environment.GetEnvironmentVariable("IMAPDEMOUSER");
                string password = Environment.GetEnvironmentVariable("IMAPDEMOPWD");
                if (user == null || password == null  ) {
                    Console.WriteLine(user);
                    Console.WriteLine("Usage: missing env vars IMAPDEMOUSER or  IMAPDEMOPWD");
                } else {

                // Create app password in Google account
                // https://support.google.com/accounts/answer/185833?hl=en
                // Gmail IMAP4 server is "imap.gmail.com"
                MailServer oServer = new MailServer("imap.gmail.com",
                                user,
                                password,
                                ServerProtocol.Imap4);

                // Enable SSL connection.
                oServer.SSLConnection = true;

                // Set 993 SSL port
                oServer.Port = 993;

                MailClient oClient = new MailClient("TryIt");
                oClient.Connect(oServer);

                // retrieve unread/new email only
                oClient.GetMailInfosParam.Reset();
                oClient.GetMailInfosParam.GetMailInfosOptions = GetMailInfosOptionType.NewOnly;

                MailInfo[] infos = oClient.GetMailInfos();
                Console.WriteLine("Total {0} unread email(s)\r\n", infos.Length);
                for (int i = 0; i < infos.Length; i++)
                {
                    MailInfo info = infos[i];
                    Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}",
                        info.Index, info.Size, info.UIDL);

                    // Receive email from IMAP4 server
                    Mail oMail = oClient.GetMail(info);

                    Console.WriteLine("From: {0}", oMail.From.ToString());
                    Console.WriteLine("Subject: {0}\r\n", oMail.Subject);

                    // Generate an unqiue email file name based on date time.
                    string fileName = _generateFileName(i + 1);
                    string fullPath = string.Format("{0}\\{1}", localInbox, fileName);

                    // Save email to local disk
                    oMail.SaveAs(fullPath, true);

                    // mark unread email as read, next time this email won't be retrieved again
                    if (!info.Read)
                    {
                        oClient.MarkAsRead(info, true);
                    }

                    // if you don't want to leave a copy on server, please use
                    // oClient.Delete(info);
                    // instead of MarkAsRead
                }

                // Quit and expunge emails marked as deleted from IMAP4 server.
                oClient.Quit();
                Console.WriteLine("Completed!");
                }
            }
            catch (Exception ep)
            {
                Console.WriteLine(ep.Message);
            }
        }
    }
}