using System;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

namespace MailSender
{
    class Program
    {
        private static void ShowHelp()
        {
            Console.WriteLine($"\n-silent\t\t\t\tLaunch program without any GUI and output, excluding log.\n" + //TODO: htmlFolderPath
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
                                    $"\nemailRecievers=\"i.ivanov@mail.com, p.petrov@mail.com\"\n");
        }

        private static void ShowConfig() //TODO: if empty - output "unconfigured"
        {
            foreach (var line in File.ReadAllLines(Configs.ConfigsPath))
            {
                Console.WriteLine(line);
            }
        }

        static void Main(string[] args) //TODO: Learn how to use catch (Exception e) and throw new exception. AND USE IT!
        {
            Logs.AddLogsCollected($"\n\n\nCurrent user: {Environment.UserName}");
            if (args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i].StartsWith("-"))
                    {
                        switch (args[i].ToLower())
                        {
                            case "-run":
                                break;
                            case "-silent":
                                break;
                            case "-help":
                                ShowHelp();
                                break;
                            case "-showconfig": //TODO: show config.cfg destination.
                                ShowConfig();
                                break;
                            case "-editconfig":
                                if (i + 1 <= args.Length)
                                {
                                    try
                                    {
                                        Configs.LoadConfig();
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
                                            Configs.EditConfig(parameter, value);
                                        }
                                        catch
                                        {
                                            Console.WriteLine("Unable to edit configuration. Invalid parameter.");
                                        }
                                        i++;
                                    }
                                }
                                Configs.SaveConfig();
                                break;
                            default:
                                Console.WriteLine("Unknown parameter.");
                                break;
                        }
                    }
                }
                if(args.Contains("-run"))
                {
                    //ADDLOGS COLLECTED!
                    //TODO: greeting string, show config and ask if everything is ok
                    Console.WriteLine($"Hello {Environment.UserName}!\n\n" +
                        $"I'm BirthdayMailSender!\n" +
                        $"There are arguments I can be run with:\n" +
                        $"-silent\t\t\tSend mail without any output using current configuration.\n" +
                        $"-help\t\t\tDisplay detailed instruction about every argument and show some usage examples.\n" +
                        $"-showconfig\t\tDisplay current configuration parameters and values, also destination of config.cfg file.\n" +
                        $"-editconig\t\tAllow user to change some parameter values.\n\n");
                    Console.WriteLine($"Here is my current configuration:\n");
                    ShowConfig();
                    Console.Write("Do you want to send message with current configuration(y/n)? ");
                    switch (Console.ReadLine().ToLower())
                    {
                        case "y":
                        case "yes":
                            try
                            {
                                Console.WriteLine("Prepare to sending...");
                                Configs.LoadConfig();
                                Sending.SendMail();
                                Logs.SendLogs();
                                Console.WriteLine("Sending succesful!");
                            }
                            catch
                            {
                                Console.WriteLine("Unable to send message. Check configuration.");
                            }
                            break;
                        case "n":
                        case "no":
                            Console.WriteLine($"Sending cancelled.");
                            break;
                        default:
                            Console.WriteLine($"It is not look's like \"yes\". Sending cancelled.");
                            break;
                    }
                }
                else if (args.Contains("-silent"))
                {
                    try
                    {
                        Logs.AddLogsCollected("Working mode: silent");
                        Configs.LoadConfig();                        
                        Sending.SendMail();
                        Logs.SendLogs();
                    }
                    catch
                    {
                        Console.WriteLine("Unable to send message. Check configuration.");
                    }
                }
            }
            else
            {
                ShowHelp();
            }
        }
    }
}
