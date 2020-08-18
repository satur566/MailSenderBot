using System;
using System.IO;

namespace MailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            Configs.AddLogsCollected($"\n\n\nCurrent user: {Environment.UserName}"); //TODO: editConfig with args -edit parameter1=value1 parameter2=value2   //TODO: -nopassword parameter //-silent parameter
            if (args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i].StartsWith("-"))
                    {
                        switch (args[i].ToLower()) {
                            case "-silent":
                                try
                                {
                                    Configs.AddLogsCollected("Working mode: silent");
                                    Methods.LoadConfig();
                                    Methods.SendMail();                                   
                                }
                                catch
                                {
                                    Console.WriteLine("Unable to send message. Check configuration.");                                    
                                }
                                break;
                            case "-help":
                                //TODO: describe every parameter.
                                break;
                            case "-showconfigure":                                
                                Configs.GetSenderEmail(); //TODO: add some console output.
                                Configs.GetSenderPassword();
                                Configs.GetSenderName();
                                Configs.GetEmailrecievers();
                                Configs.GetMessageSubject();
                                Configs.GetHtmlPath();
                                Configs.GetXlsPath();
                                Configs.GetEmployeeNameColumnNumber();
                                Configs.GetBirthdayColumnNumber();
                                Configs.GetServerAddress();
                                Configs.GetServerPort();
                                Configs.GetFiveDayMode();
                                Configs.GetLogRecievers();
                                break;
                            case "-editconfig": //TODO: edit one or more parameters until '-' occured.
                                break;
                            default:
                                Console.WriteLine("Unknown parameter.");
                                break;
                        }
                    }
                }
            }                        
        }
    }
}
