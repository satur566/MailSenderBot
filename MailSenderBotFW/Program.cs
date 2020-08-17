using System;
using System.IO;

namespace MailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            Configs.AddLogsCollected($"\n\n\nCurrent user: {Environment.UserName}"); //TODO: editConfig with args -edit parameter1=value1 parameter2=value2   
            try
            {
                Methods.LoadConfig();                                
            }
            catch
            {
                Console.WriteLine("Cannot properly read config file. Please set configuration again.");
                Console.Write("Set sender e-mail: ");
                Methods.EditConfig("text", "senderEmail", Console.ReadLine());

                Console.Write("Set sender password: ");
                Methods.EditConfig("password", "senderPassword", Console.ReadLine()); //TODO: password enter with mask.

                Console.Write("Set sender displayed name: ");
                Methods.EditConfig("text", "senderName", Console.ReadLine());

                Console.Write("Set reciever e-mail: ");
                Methods.EditConfig("text", "recieverEmail", Console.ReadLine());

                Console.Write("Set message subject: ");
                Methods.EditConfig("text", "messageSubject", Console.ReadLine());

                Console.Write($"Set a path to html file: ");
                Methods.EditConfig("html", "htmlPath", Console.ReadLine());              

                Console.Write($"Set a path to xls file: ");
                Methods.EditConfig("xls", "xlsPath", Console.ReadLine());

                Console.Write("Set a number of column contains birthday dates: ");
                Methods.EditConfig("digit", "birthdayColumnNumber", Console.ReadLine());           

                Console.Write("Set a number of column contains employees names: ");
                Methods.EditConfig("digit", "employeeNameColumnNumber", Console.ReadLine()); 

                Console.Write("Set server address: ");
                Methods.EditConfig("text", "serverAddress", Console.ReadLine());

                Console.Write("Set server port (if default - leave empty): ");
                Methods.EditConfig("port", "serverPort", Console.ReadLine());   

                Console.WriteLine("Use 5/2 workmode?(yes) \nOtherwise will be user full week mode");
                Methods.EditConfig("y/n", "fiveDaysMode", Console.ReadLine());

                Console.Write("Set logs recievers: ");
                Methods.EditConfig("text", "logRecievers", Console.ReadLine()); //TODO: ConfigWriter remade for this.               
                try
                {
                    File.WriteAllText(Configs.GetConfigPath(), string.Empty);
                    File.WriteAllLines(Configs.GetConfigPath(), Configs.GetConfigurations().ToArray());
                    Configs.AddLogsCollected($"Config save: SUCCESS.");
                }
                catch
                {
                    Configs.AddLogsCollected($"Config save: FAILURE.");
                }
                Methods.LoadConfig();              
                Configs.SetReadConfigSuccess(false);
            }
            /*try
            {
                if (DateTime.Now.Day == 2 && DateTime.Now.Month == 8)
                {
                    Methods.SendMessage("a.maksimov@sever.ttk.ru", "Happy Birthday!", "Happy birthday, daddy! Wish you a good incoming year!", args, false);
                    Methods.SendMessage("satur566@gmail.com", "Happy Birthday!", "Happy birthday, daddy! Wish you a good incoming year!", args, false);
                }
            }
            catch { }*/
            Methods.SendMail();
        }       
    }
}
