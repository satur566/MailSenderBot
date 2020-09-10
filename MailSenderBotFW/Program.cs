using System;
using System.IO;
using System.Linq;

namespace MailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            Logs.AddLogsCollected($"Current user: {Environment.UserName}");
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
                            case "-showconfig":
                                ShowConfig();
                                break;
                            case "-editconfig":
                                EditConfiguration(ref i, args);
                                break;
                            default:
                                Console.WriteLine("Unknown parameter.");
                                break;
                        }
                    }
                }
                if (args.Contains("-run"))
                {
                    RunApp();
                }
                else if (args.Contains("-silent"))
                {
                    Silent();
                }
            }
            else
            {
                ShowHelp();
            }
        }

        private static void RunApp()
        {
            Logs.AddLogsCollected("Working mode: dialogue");
            //TODO: ADDLOGS COLLECTED!
            Console.WriteLine($"Hello {Environment.UserName}!\n\n" +
                $"I'm BirthdayMailSender!\n" +
                $"There are arguments I can be run with:\n" +
                $"-run\t\t\tRuns the program in dialogue mode with user." +
                $"-silent\t\t\t\tLaunch program without any GUI and output, excluding log.\n" +
                $"-help\t\t\t\tDisplays help.\n" +
                $"-showconfig\t\t\tShow current configuration stored in config.cfg file.\n" +
                $"-editconfig\t\t\tEdit current configuration stored in config.cfg file.");
            Console.WriteLine($"Here is my current configuration:\n");
            ShowConfig();
            Console.Write("Do you want to send message with current configuration(y/n)? ");
            switch (Console.ReadLine().ToLower())
            {
                case "y":
                case "yes":
                    Console.WriteLine("Prepare to sending...");
                    SendAll();
                    break;
                case "n":
                case "no":
                    Logs.AddLogsCollected("Sending message cancelled.");
                    Logs.AddLogsCollected("Reason: user cancel.");
                    Console.WriteLine($"Sending cancelled.");
                    break;
                default:
                    Console.WriteLine($"It is not look's like \"yes\". Sending cancelled.");
                    Logs.AddLogsCollected("Sending message cancelled.");
                    Logs.AddLogsCollected("Reason: incorrect user input.");
                    break;
            }
        }
        private static void EditConfiguration(ref int i, string[] argunemts)
        {
            if (i + 1 <= argunemts.Length)
            {
                try
                {
                    Configs.ParametersList = FileWorks.LoadConfig(FileWorks.ConfigsPath);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Unable to load previous configuration: {e.Message}");
                }
                for (int j = i + 1; j < argunemts.Length && !argunemts[j].StartsWith("-"); j++)
                {
                    try
                    {
                        string parameter = argunemts[j].Substring(0, argunemts[j].IndexOf('='));
                        string value = argunemts[j].Substring(argunemts[j].IndexOf('=') + 1, argunemts[j].Length - argunemts[j].IndexOf('=') - 1);
                        Configs.EditConfig(parameter, value);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"Unable to edit configuration: {e.Message}");
                    }
                    i++;
                }
            }
            FileWorks.SaveConfig();
        }
        private static void Silent()
        {
            Logs.AddLogsCollected("Working mode: silent");            
            SendAll();
        }
        private static void ShowHelp()
        {
            Console.WriteLine($"\n-run\t\t\t\tRuns the program in dialogue mode with user.\n" +
                                    "-silent\t\t\t\tLaunch program without any GUI and output, excluding log.\n" +
                                    $"-help\t\t\t\tDisplays help.\n" +
                                    $"-showconfig\t\t\tShow current configuration stored in config.cfg file.\n" +                                    
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
                                    $"htmlFolderPath\t\t\tPath to folder which contains at least one html file.\n" +
                                    $"htmlSwitchMode\t\t\tSwitch mode of message body:\n" +
                                    $"\t\t\t\t\trandom - pick message body from random html-sample file each launch.\n" +
                                    $"\t\t\t\t\tascending - pick message body in ascendingly ordered html-sample \n" +
                                    $"\t\t\t\t\tfile each launch.\n" +
                                    $"\t\t\t\t\tno switch - disable pick new body.\n" + 
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
        private static void ShowConfig()
        {
            Console.WriteLine($"Configuration file name: {FileWorks.ConfigsPath}\n" +
                $"\nContent: ");
            string[] configFileContent = File.ReadAllLines(FileWorks.ConfigsPath);
            if (configFileContent.Length.Equals(0))
            {
                Console.Write("nothing to display, file is empty.");
            }
            foreach (string line in configFileContent)
            {
                Console.WriteLine(line);
            }
        }     
        
        private static void SendAll()
        {
            Configs.ParametersList = FileWorks.LoadConfig(FileWorks.ConfigsPath);
            try
            {
                Sending.SendMail();
                Console.WriteLine("Sending messages: SUCCESS.");
            }
            catch (Exception e)
            {
                string exceptionMessage = $"Sending messages: FAILURE.\n\t\tReason: {e.Message}";
                Console.WriteLine(exceptionMessage);
                Logs.AddLogsCollected(exceptionMessage);
            }
            try
            {
                Logs.SendLogs();
                Console.WriteLine("Sending logs: SUCCESS.");
            }
            catch (Exception e)
            {
                string exceptionMessage = $"Sending logs: FAILURE.\n\t\tReason: {e.Message}";
                Console.WriteLine(exceptionMessage);
                Logs.AddLogsCollected(exceptionMessage);
            }
        }
    }
}
