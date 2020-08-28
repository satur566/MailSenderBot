using System;
using System.IO;
using System.Linq;

namespace MailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            Configs.AddLogsCollected($"\n\n\nCurrent user: {Environment.UserName}");
            if (args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i].StartsWith("-"))
                    {
                        switch (args[i].ToLower())
                        {
                            case "-silent":
                                break;
                            case "-help":
                                Console.WriteLine($"\n-silent\t\t\t\tLaunch program without any GUI and output, excluding log.\n" +
                                    $"-showconfig\t\t\tShow current configuration stored in config.cfg file.\n" +
                                    $"-help\t\t\t\tDisplays help.\n" +
                                    $"-editconfig\t\t\tEdit current configuration stored in config.cfg file. " +
                                    $"\n\t\t\t\tUsage: -editconfig parameter=\"value\" \n" +                                    
                                    $"\nList of parameters available:\n\n" +
                                    $"senderEmail\t\t\tE-mail address of sender.\n" +
                                    $"senderUsername\t\t\tE-mail server authorisation username.\n" +
                                    $"senderPassword\t\t\tE-mail server authorisation username. Note that you cannot change password via" +
                                    $"\n\t\t\t\tconfig.cfg manually due to encoding/decoding.\n" +
                                    $"senderName\t\t\tSender name displayed to reciever.\n\n" +
                                    $"emailRecievers\t\t\tList of e-mail recievers comma separated.\n" +
                                    $"messageSubject\t\t\tDisplayed subject of e-mail.\n" +
                                    $"htmlPath\t\t\tPath to html file. File should contain at least %LIST_OF_EMPLOYEES% string " +
                                    $"\n\t\t\t\tinside and has .html file extension.\n" +
                                    $"xlsPath\t\t\t\tPath to xls file. File should has .xls file extension.\n" +
                                    $"birthdayColumnNumber\t\tNumber of column, contains date of employee birthday. " +
                                    $"\n\t\t\t\tNote that date format starts with \' i.e \'{DateTime.Now}\n" +
                                    $"employeeNameColumnNumber\tNumber of column, contains employee full name.\n" +
                                    $"serverAddress\t\t\tIP-address of e-mail server.\n" +
                                    $"serverPort\t\t\tPort of e-mail server. If leaved empty - used default 25 port.\n" +
                                    $"fiveDaysMode\t\t\tWorking mode that allow send e-mails only from Monday to Friday. " +
                                    $"\n\t\t\t\tEmployees which birthday date happened on Saturday or Sunday " +
                                    $"\n\t\t\t\twill be congratulated on Monday.\n" +
                                    $"logRecievers\t\t\tList of logs recievers comma separated.\n" +
                                    $"\nUsage exaple:\n\n" +
                                    $"-editconfig senderEmail=info@mail.com senderPassword=Qwerty123 htmlPath=C:\\temp\\file.html " +
                                    $"\nemailRecievers=i.ivanov@mail.com, p.petrov@mail.com\n");
                                break;
                            case "-showconfig":
                                foreach (var line in File.ReadAllLines(Configs.ConfigsPath))
                                {
                                    Console.WriteLine(line);
                                }
                                break;
                            case "-editconfig":
                                if (i + 1 <= args.Length)
                                {
                                    try
                                    {
                                        Methods.LoadConfig();
                                    }
                                    catch
                                    {
                                        Console.WriteLine("Unable to find previous configuration.");
                                    }
                                    for (int j = i + 1; j < args.Length && !args[j].StartsWith("-"); j++)
                                    {
                                        try
                                        {
                                            string parameter = args[j].Substring(0, args[j].IndexOf('='));
                                            string value = args[j].Substring(args[j].IndexOf('=') + 1, args[j].Length - args[j].IndexOf('=') - 1);
                                            Methods.EditConfig(parameter, value);
                                        }
                                        catch
                                        {
                                            Console.WriteLine("Unable to edit configuration. Invalid parameter.");
                                        }
                                        i++;
                                    }
                                }
                                Methods.SaveConfig();
                                break;
                            default:
                                Console.WriteLine("Unknown parameter.");
                                break;
                        }
                    }
                }
                if (args.Contains("-silent"))
                {
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
                }
            }
            else
            {
                //TODO: launch WPF app.
            }
        }
    }
}
