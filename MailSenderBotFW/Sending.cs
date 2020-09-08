using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;

namespace MailSender
{
    static class Sending
    {
        public static void SendMail()
        {
            if (string.IsNullOrEmpty(Configs.HtmlFilePath) || string.IsNullOrWhiteSpace(Configs.HtmlFilePath))
            {
                Configs.EditConfig("htmlPath", Configs.RandomChangeHtmlFile(Configs.HtmlFilePath));
                FileWorks.SaveConfig();
            }
            Employees.WhosBirthdayIs = FileWorks.ReadXlsFile(Configs.XlsFilePath, Configs.FiveDayMode, Configs.BirthdayColumnNumber, Configs.EmployeeNameColumnNumber);            
            Configs.MessageText = FileWorks.ReadHtmlFile(Configs.HtmlFilePath, Employees.CongratulationsList);
            Logs.AddLogsCollected($"Selected {Configs.HtmlFilePath} sample of html body.");            
            if (Employees.WhosBirthdayIs.Count.Equals(0) || Configs.FiveDayMode && (DateTime.Now.DayOfWeek == DayOfWeek.Sunday || DateTime.Now.DayOfWeek == DayOfWeek.Saturday))
            {
                Logs.AddLogsCollected($"Sending message: CANCELLED.");
                if (Configs.FiveDayMode && (DateTime.Now.DayOfWeek == DayOfWeek.Sunday || DateTime.Now.DayOfWeek == DayOfWeek.Saturday))
                {
                    Logs.AddLogsCollected($"Reason: today is a day off.");
                }
                if (Employees.WhosBirthdayIs.Count.Equals(0))
                {
                    Logs.AddLogsCollected($"Reason: employees don't have a birthday today.");
                }
                else if (Configs.FiveDayMode && (DateTime.Now.DayOfWeek == DayOfWeek.Sunday || DateTime.Now.DayOfWeek == DayOfWeek.Saturday))
                {
                    Logs.AddLogsCollected($"On Monday {Employees.WhosBirthdayIs.Count} employees will be congratulated.");
                }
            }
            else
            {
                if (Configs.EmailRecievers.Count.Equals(0))
                {
                    Logs.AddLogsCollected($"Sending message: CANCELLED.");
                    Logs.AddLogsCollected($"Reason: recievers count equals 0.");
                }
                else
                {
                    foreach (var reciever in Configs.EmailRecievers)
                    {
                        try
                        {
                            SendMessage(reciever, Configs.MessageSubject, Configs.MessageText);
                            Logs.AddLogsCollected($"Sending message to {reciever}: SUCCESS.");
                            Logs.AddLogsCollected(Logs.LogConclusionMaker(reciever));
                        } 
                        catch (Exception e)
                        {
                            throw e;
                        }
                    }
                    switch (Configs.HtmlSwitchMode.ToLower())
                    {
                        case "ascending":
                            Configs.EditConfig("htmlPath", Configs.AscendingChangeHtmlFile(Configs.HtmlFilePath));
                            FileWorks.SaveConfig();
                            break;
                        case "random":
                            Configs.EditConfig("htmlPath", Configs.RandomChangeHtmlFile(Configs.HtmlFilePath));
                            FileWorks.SaveConfig();
                            break;
                        default:
                            break;
                    }                                      
                }
            }
        }        

        public static void SendMessage(string reciever, string subject, string message) //TODO: try to make throw in get properties in Configs.sendermail etc if empty.
        {
            try
            {
                MailAddress Sender = new MailAddress(Configs.SenderEmail, Configs.SenderName);
                MailAddress Reciever = new MailAddress(reciever);
                MailMessage Message = new MailMessage(Sender, Reciever)
                {
                    Subject = subject,
                    Body = message,
                    IsBodyHtml = true
                };
                List<LinkedResource> images = new List<LinkedResource>();
                string[] htmlArray = message.Split(new string[] { System.Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                string htmlFolderPath = Configs.HtmlFilePath.Substring(0, Configs.HtmlFilePath.LastIndexOf('\\') + 1);
                int imageCounter = 0;
                foreach (var line in htmlArray) //Try catch if no images in path selected in src in html
                {
                    if (line.Contains("src="))
                    {
                        try
                        {
                            string srcLine = line.Substring(line.IndexOf('\"') + 1, line.Substring(line.IndexOf('\"') + 1).IndexOf('\"'));
                            string imagePath = htmlFolderPath + srcLine.Replace('/', '\\');
                            images.Add(new LinkedResource(imagePath, "image/gif"));
                            message = message.Replace(srcLine, "cid:" + images[imageCounter++].ContentId);
                        }
                        catch
                        {
                            throw new Exception("Unable to attach images from html file.");
                        }
                    }
                }
                var htmlView = AlternateView.CreateAlternateViewFromString(message, Encoding.UTF8, MediaTypeNames.Text.Html);
                images.ForEach(htmlView.LinkedResources.Add);
                Message.AlternateViews.Add(htmlView);
                SmtpClient Client = new SmtpClient(Configs.ServerAddress, Configs.ServerPort)
                {
                    Credentials = new NetworkCredential(Configs.SenderUsername, Encryptor.DecryptString("b14ca5898a4e4133bbce2mbd02082020", Configs.SenderPassword)),
                    EnableSsl = false
                };
                try
                {
                    Client.Send(Message);
                }
                catch (Exception e)
                {
                    throw new Exception($"Unexpected error happened. {e.Message}");
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
