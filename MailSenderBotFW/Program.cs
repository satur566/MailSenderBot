using System;
using System.IO;

namespace MailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            Configs.AddLogsCollected($"\n\n\nCurrent user: {Environment.UserName}");            
            try
            {
                Methods.ReadConfig();                                
            }
            catch
            {
                Console.WriteLine("Cannot properly read config file. Please set configuration again.");
                Console.Write("Set sender e-mail: ");
                Methods.ConfigWriter("text", "senderEmail", Console.ReadLine());

                Console.Write("Set sender password: ");
                Methods.ConfigWriter("password", "senderPassword", Console.ReadLine()); //TODO: password enter with mask.

                Console.Write("Set sender displayed name: ");
                Methods.ConfigWriter("text", "senderName", Console.ReadLine());

                Console.Write("Set reciever e-mail: ");
                Methods.ConfigWriter("text", "recieverEmail", Console.ReadLine());

                Console.Write("Set message subject: ");
                Methods.ConfigWriter("text", "messageSubject", Console.ReadLine());

                Console.Write($"Set a path to html file: ");
                Methods.ConfigWriter("html", "htmlPath", Console.ReadLine()); //Remade IsFileExist                

                Console.Write($"Set a path to xls file: ");
                Methods.ConfigWriter("xls", "xlsPath", Console.ReadLine()); //Remade IsFileExist

                Console.Write("Set a number of column contains birthday dates: ");
                Methods.ConfigWriter("digit", "birthdayColumnNumber", Console.ReadLine()); //Remade IsDigit            

                Console.Write("Set a number of column contains employees names: ");
                Methods.ConfigWriter("digit", "employeeNameColumnNumber", Console.ReadLine()); //Remade IsDigit    

                Console.Write("Set server address: ");
                Methods.ConfigWriter("text", "serverAddress", Console.ReadLine());

                Console.Write("Set server port (if default - leave empty): ");
                Methods.ConfigWriter("port", "serverPort", Console.ReadLine()); //Remade IsDigit     

                Console.WriteLine("Use 5/2 workmode?(yes) \nOtherwise will be user full week mode");
                Methods.ConfigWriter("text", "fiveDaysMode", Console.ReadLine()); //TODO: ConfigWriter remade for this. (yes/no)

                Console.Write("Set logs recievers: ");
                Methods.ConfigWriter("text", "logRecievers", Console.ReadLine()); //TODO: ConfigWriter remade for this.
                /*
                string recieversString = Console.ReadLine();
                string[] recievers = recieversString.Split(',');
                foreach (var reciever in recievers)
                {
                    Configs.SetLogRecievers(reciever.Trim());
                }
                */
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
                Methods.ReadXlsFile(Configs.GetXlsPath(), Configs.GetFiveDayMode(), Configs.GetBirthdayColumnNumber(), Configs.GetEmployeeNameColumnNumber());
                Methods.ReadHtmlFile(Configs.GetHtmlPath(), Employees.GetCongratulationsString());
                Methods.ReadConfig();
                Configs.SetReadConfigSuccess(false);
            }
            if (Configs.GetReadConfigSuccess())
            {
                Configs.AddLogsCollected($"Loading config: SUCCESS");
            } 
            else
            {
                Configs.AddLogsCollected($"Loading config: FAILURE");
            }
            try
            {
                /*if (DateTime.Now.Day == 2 && DateTime.Now.Month == 8)
                {
                    Methods.SendMessage("a.maksimov@sever.ttk.ru", "Happy Birthday!", "Happy birthday, daddy! Wish you a good incoming year!", args, false);
                    Methods.SendMessage("satur566@gmail.com", "Happy Birthday!", "Happy birthday, daddy! Wish you a good incoming year!", args, false);
                }*/
            }
            catch { }
            Methods.SendMail(args);
        }       
    }
}
