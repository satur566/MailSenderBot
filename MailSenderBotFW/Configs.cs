using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace MailSender
{
    static class Configs
    {
        private static string senderEmail;
        private static string senderName;
        private static string senderUsername;
        private static string senderPassword;
        private static List<string> emailRecievers = new List<string>();
        private static string messageSubject;
        private static string messageText;
        private static string serverAddress;
        private static string serverPort;
        private static List<string> configurations = new List<string>();
        private static string htmlPath;
        private static string xlsPath;
        private static bool fiveDayMode;
        private static string birthdayColumnNumber;
        private static string employeeNameColumnNumber;
        private static readonly List<string> logRecievers = new List<string>();
        private static string logsCollected = "";
        //TODO: switch html samples like: every day, every week, every month, every time of the year.
        //TODO: switch should be random but never same in a row twice and ascendingly ordered by name of html file.

        public static string GetSenderEmail()
        {
            return senderEmail;
        }
        public static void SetSenderEmail(string email)
        {
            senderEmail = email;
        }

        public static string GetSenderName()
        {
            return senderName;
        }
        public static void SetSenderName(string name)
        {
            senderName = name;
        }

        public static string GetSenderUsername()
        {
            return senderUsername;
        }
        public static void SetSenderUsername(string username)
        {
            senderUsername = username;
        }

        public static string GetSenderPassword()
        {
            return senderPassword;
        }
        public static void SetSenderPassword(string password)
        {
            senderPassword = password;
        }

        public static List<string> GetEmailrecievers()
        {
            return emailRecievers;
        }
        public static void SetEmailRecievers(List<string> recievers)
        {
            emailRecievers.Clear();
            foreach (var reciever in recievers)
            {
                emailRecievers.Add(reciever);
            }
        }

        public static string GetMessageSubject()
        {
            return messageSubject;
        }
        public static void SetMessageSubject(string subject)
        {
            messageSubject = subject;
        }

        public static string GetMessageText()
        {
            return messageText;
        }
        public static void SetMessageText(string text)
        {
            messageText = text;
        }

        public static string GetServerAddress()
        {
            return serverAddress;
        }
        public static void SetServerAddress(string address)
        {
            serverAddress = address;
        }

        public static string GetServerPort()
        {
            return serverPort;
        }
        public static void SetServerPort(string port)
        {
            serverPort = port;
        }

        public static string GetLogsPath()
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

        public static string GetConfigPath()
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

        private static void Move<T>(this List<T> list, int oldIndex, int newIndex)
        {
            T item = list[oldIndex];
            list.RemoveAt(oldIndex);
            list.Insert(newIndex, item);
        }

        private static void SortConfiguration(ref List<string> list, string parameter)
        {
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains(parameter))))
            {
                list.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains(parameter)))]);
            }
        }
        public static List<string> GetConfigurations() //TODO: think about much more simple sorting method.
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
            /*
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains("senderEmail"))))
            {
                tempList.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains("senderEmail")))]);
            }
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains("senderUsername"))))
            {
                tempList.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains("senderUsername")))]);
            }
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains("senderPassword"))))
            {
                tempList.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains("senderPassword")))]);
            }
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains("senderName"))))
            {
                tempList.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains("senderName")))]);
            }
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains("emailRecievers"))))
            {
                tempList.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains("emailRecievers")))]);
            }
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains("messageSubject"))))
            {
                tempList.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains("messageSubject")))]);
            }
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains("htmlPath"))))
            {
                tempList.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains("htmlPath")))]);
            }
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains("xlsPath"))))
            {
                tempList.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains("xlsPath")))]);
            }
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains("birthdayColumnNumber"))))
            {
                tempList.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains("birthdayColumnNumber")))]);
            }
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains("employeeNameColumnNumber"))))
            {
                tempList.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains("employeeNameColumnNumber")))]);
            }
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains("serverAddress"))))
            {
                tempList.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains("serverAddress")))]);
            }
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains("serverPort"))))
            {
                tempList.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains("serverPort")))]);
            }
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains("fiveDaysMode"))))
            {
                tempList.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains("fiveDaysMode")))]);
            }
            if (configurations.Contains(configurations.FirstOrDefault(value => value.Contains("logRecievers"))))
            {
                tempList.Add(configurations[configurations.IndexOf(configurations.FirstOrDefault(value => value.Contains("logRecievers")))]);
            }*/
            configurations = new List<string>(tempList);
            return configurations;
        }
        public static void SetConfigurations(List<string> configList)
        {
            configurations = configList;
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

        public static string GetHtmlPath()
        {
            return htmlPath;
        }
        public static void SetHtmlPath(string path)
        {
            htmlPath = path;
        }

        public static string GetXlsPath()
        {
            return xlsPath;
        }
        public static void SetXlsPath(string path)
        {
            xlsPath = path;
        }

        public static bool GetFiveDayMode()
        {
            return fiveDayMode;
        }
        public static void SetFiveDayMode(bool state)
        {
            fiveDayMode = state;
        }

        public static string GetBirthdayColumnNumber()
        {
            return birthdayColumnNumber;
        }
        public static void SetBirthdayColumnNumber(string number)
        {
            birthdayColumnNumber = number;
        }

        public static string GetEmployeeNameColumnNumber()
        {
            return employeeNameColumnNumber;
        }
        public static void SetEmployeeNameColumnNumber(string number)
        {
            employeeNameColumnNumber = number;
        }

        public static string GetLogsCollected()
        {
            return logsCollected;
        }
        public static void AddLogsCollected(string log)
        {
            log = $"\n{DateTime.Now} - " + log;
            File.AppendAllText(Configs.GetLogsPath(), log); //Try.
            logsCollected = String.Concat(logsCollected, log.Replace("\t", "&#9;").Replace("\n", "<br>"));
        }

        public static List<string> GetLogRecievers()
        {
            return logRecievers;
        }
        public static void SetLogRecievers(List<string> recievers)
        {
            logRecievers.Clear();
            foreach (var reciever in recievers)
            {
                logRecievers.Add(reciever);
            }
        }
    }
}