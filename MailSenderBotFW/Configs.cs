using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Policy;

namespace MailSender
{
    static class Configs
    {
        private static List<string> emailRecievers = new List<string>();
        private static List<string> configurations = new List<string>();
        private static string xlsFilePath;
        private static readonly List<string> logRecievers = new List<string>();
        private static string logsCollected = "";
        //TODO: switch html samples like: every day, every week, every month, every time of the year.
        //TODO: switch should be random but never same in a row twice and ascendingly ordered by name of html file.

        public static string SenderEmail { set; get; }

        public static string SenderName { set; get; }

        public static string SenderUsername { set; get; }

        public static string SenderPassword { set; get; }

        public static List<string> EmailRecievers
        {
            get
            {
                return emailRecievers;
            }
            set
            {
                emailRecievers.Clear();
                foreach (string reciever in value)
                {
                    emailRecievers.Add(reciever);
                }
            }
        }

        public static string MessageSubject { set; get; }

        public static string MessageText { set; get; }

        public static string ServerAddress { set; get; }

        public static int ServerPort { set; get; }        

        public static string LogsPath
        {
            get
            {
                string logsDirectory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\logs";
                string logFile = $"\\{DateTime.Now.Month:D}-{DateTime.Now.Year}.log";
                if (!File.Exists(logsDirectory + logFile))
                {
                    try
                    {
                        Directory.CreateDirectory(logsDirectory);
                        var file = File.Create(logsDirectory + logFile);
                        file.Close();
                    }
                    catch
                    {
                        Console.WriteLine("Cannot create log file in readOnly directory. Press any key to exit");
                    }

                }
                return logsDirectory + logFile;
            }
        }

        public static string ConfigsPath
        {
            get
            {
                string configDirectory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\etc";
                if (!File.Exists(configDirectory + "\\config.cfg"))
                {
                    try
                    {
                        Directory.CreateDirectory(configDirectory);
                        var file = File.Create(configDirectory + "\\config.cfg");
                        file.Close();
                        File.WriteAllLines(configDirectory + "\\config.cfg", configurations);
                    }
                    catch
                    {
                        Console.WriteLine("Cannot create config file in readOnly directory. Press any key to exit");
                        Environment.Exit(0);
                    }

                }
                return configDirectory + "\\config.cfg";
            }
        }

        private static void SortConfiguration(ref List<string> list, string parameter)
        {
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains(parameter))))
            {
                list.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains(parameter)))]);
            }
        }

        public static List<string> ConfigurationsList
        {
            get
            {
                List<string> tempList = new List<string>();
                SortConfiguration(ref tempList, "senderEmail");
                SortConfiguration(ref tempList, "senderUsername");
                SortConfiguration(ref tempList, "senderPassword");
                SortConfiguration(ref tempList, "senderName");
                SortConfiguration(ref tempList, "emailRecievers");
                SortConfiguration(ref tempList, "messageSubject");
                SortConfiguration(ref tempList, "htmlPath"); //TODO: if folder - check available file inside, i.e contains listofusers and ends with .html
                SortConfiguration(ref tempList, "xlsPath");
                SortConfiguration(ref tempList, "birthdayColumnNumber");
                SortConfiguration(ref tempList, "employeeNameColumnNumber");
                SortConfiguration(ref tempList, "serverAddress");
                SortConfiguration(ref tempList, "serverPort");
                SortConfiguration(ref tempList, "fiveDaysMode");
                SortConfiguration(ref tempList, "logRecievers");
                configurations = new List<string>(tempList);
                return configurations;
            }
            set
            {
                configurations = value;
            }
        }

        public static void ChangeConfigurations(string parameter, string value)
        {            
            if (configurations.Contains(configurations.FirstOrDefault(item => item.Contains(parameter))))
            {
                configurations.Remove(configurations.FirstOrDefault(item => item.Contains(parameter)));
                configurations.Add(parameter + "=" + value);
                Configs.AddLogsCollected($"Config changed: " + parameter + "=" + value);
            }
            else
            {
                configurations.Add(parameter + "=" + value);
                Configs.AddLogsCollected($"Config added: " + parameter + "=" + value);
            }
        }

        public static string HtmlFilePath { get; set; }

        public static string XlsFilePath
        {
            get
            {
                return xlsFilePath;
            }
            set
            {
                xlsFilePath = value;
            }
        }

        public static bool FiveDayMode { get; set; }

        public static string BirthdayColumnNumber { get; set; }
        public static string EmployeeNameColumnNumber { get; set; }

        public static string LogsCollected
        {
            get
            {
                return logsCollected;
            }
        }

        public static void AddLogsCollected(string log)
        {
            log = $"\n{DateTime.Now} - " + log;
            File.AppendAllText(Configs.LogsPath, log);
            logsCollected = String.Concat(logsCollected, log.Replace("\t", "&#9;").Replace("\n", "<br>"));
        }

        public static List<string> LogsRecievers
        {
            get
            {
                return logRecievers;
            }
            set
            {
                logRecievers.Clear();
                foreach (var reciever in value)
                {
                    logRecievers.Add(reciever);
                }
            }
        }
    }
}