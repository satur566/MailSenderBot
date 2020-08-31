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
        private static readonly List<string> emailRecievers = new List<string>();
        private static List<string> parametersList = new List<string>();
        private static readonly List<string> logRecievers = new List<string>();
        private static readonly string workingDirectory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

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
                string logsDirectory = workingDirectory + "\\logs";
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
                        AddLogsCollected("Unable to create logs directory.");
                    }

                }
                return logsDirectory + logFile;
            }
        }

        public static string ConfigsPath
        {
            get
            {
                string configDirectory = workingDirectory + "\\etc";
                if (!File.Exists(configDirectory + "\\config.cfg"))
                {
                    try
                    {
                        Directory.CreateDirectory(configDirectory);
                        var file = File.Create(configDirectory + "\\config.cfg");
                        file.Close();
                        File.WriteAllLines(configDirectory + "\\config.cfg", parametersList);
                    }
                    catch
                    {
                        AddLogsCollected("Unable to create config directory.");
                    }

                }
                return configDirectory + "\\config.cfg";
            }
        }

        public static string TempPath
        {
            get
            {
                string tempDirectory = workingDirectory + "\\temp";
                if(!Directory.Exists(tempDirectory))
                {
                    try
                    {
                        Directory.CreateDirectory(tempDirectory);
                    }
                    catch
                    {
                        AddLogsCollected("Unable to create temp directory");
                    }

                }
                return tempDirectory;
            }
        }

        private static void SortConfiguration(ref List<string> list, string parameter)
        {
            if (parametersList.Contains(parametersList.FirstOrDefault(value => value.Contains(parameter))))
            {
                list.Add(parametersList[parametersList.IndexOf(parametersList.FirstOrDefault(value => value.Contains(parameter)))]);
            }
        }

        public static List<string> ParametersList
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
                parametersList = new List<string>(tempList);
                return parametersList;
            }
            set
            {
                parametersList = value;
            }
        }

        public static void ChangeParameter(string parameter, string value)
        {            
            if (parametersList.Contains(parametersList.FirstOrDefault(item => item.Contains(parameter))))
            {
                parametersList.Remove(parametersList.FirstOrDefault(item => item.Contains(parameter)));
                parametersList.Add(parameter + "=" + value);
                Configs.AddLogsCollected($"Config changed: " + parameter + "=" + value);
            }
            else
            {
                parametersList.Add(parameter + "=" + value);
                Configs.AddLogsCollected($"Config added: " + parameter + "=" + value);
            }
        }

        public static string HtmlFilePath { get; set; }

        public static string XlsFilePath { get; set; }

        public static bool FiveDayMode { get; set; }

        public static string BirthdayColumnNumber { get; set; }
        public static string EmployeeNameColumnNumber { get; set; }

        public static string LogsCollected { get; private set; } = "";

        public static void AddLogsCollected(string log)
        {
            log = $"\n{DateTime.Now} - " + log;
            File.AppendAllText(Configs.LogsPath, log);
            LogsCollected = String.Concat(LogsCollected, log.Replace("\t", "&#9;").Replace("\n", "<br>"));
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