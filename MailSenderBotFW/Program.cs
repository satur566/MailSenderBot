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
                                break;
                            case "-showconfigure":
                                break;


                            default:
                                Console.WriteLine("Unknown parameter.");
                                break;
                        }
                    }
                }
            }            
            /*try
            {
                Methods.LoadConfig();
            }
            catch
            {
                Console.WriteLine("Cannot properly read config file. Please set configuration again.");
                Console.Write("Set sender e-mail: ");
                Methods.EditConfig("senderEmail", Console.ReadLine());

                Console.Write("Set sender password: ");
                Methods.EditConfig("senderPassword", Console.ReadLine()); //TODO: password enter with mask low priority.

                Console.Write("Set sender displayed name: "); //TODO: return username
                Methods.EditConfig("senderName", Console.ReadLine());

                Console.Write("Set reciever e-mail: ");
                Methods.EditConfig("emailRecievers", Console.ReadLine());

                Console.Write("Set message subject: ");
                Methods.EditConfig("messageSubject", Console.ReadLine());

                Console.Write($"Set a path to html file: ");
                Methods.EditConfig("htmlPath", Console.ReadLine());

                Console.Write($"Set a path to xls file: ");
                Methods.EditConfig("xlsPath", Console.ReadLine());

                Console.Write("Set a number of column contains birthday dates: ");
                Methods.EditConfig("birthdayColumnNumber", Console.ReadLine());

                Console.Write("Set a number of column contains employees names: ");
                Methods.EditConfig("employeeNameColumnNumber", Console.ReadLine());

                Console.Write("Set server address: ");
                Methods.EditConfig("serverAddress", Console.ReadLine());

                Console.Write("Set server port (if default - leave empty): ");
                Methods.EditConfig("serverPort", Console.ReadLine());

                Console.WriteLine("Use 5/2 workmode?(yes) \nOtherwise will be user full week mode");
                Methods.EditConfig("fiveDaysMode", Console.ReadLine());

                Console.Write("Set logs recievers: ");
                Methods.EditConfig("logRecievers", Console.ReadLine());
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
            }
            Methods.SendMail();*/
        }
    }
}
