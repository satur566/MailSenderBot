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

        //TODO: make replaceable string %CONGRATULATION_TEXT% similar as %LIST_OF_USERS% and reads that text from txt file.
        //TODO: make property of congratulationText path
         
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
                        Logs.AddLogsCollected("Unable to create logs directory.");
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
                        Logs.AddLogsCollected("Unable to create config directory.");
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
                        Logs.AddLogsCollected("Unable to create temp directory");
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

        private static void ChangeParameter(string parameter, string value)
        {            
            if (parametersList.Contains(parametersList.FirstOrDefault(item => item.Contains(parameter))))
            {
                parametersList.Remove(parametersList.FirstOrDefault(item => item.Contains(parameter)));
                parametersList.Add(parameter + "=" + value);
                Logs.AddLogsCollected($"Config changed: " + parameter + "=" + value);
            }
            else
            {
                parametersList.Add(parameter + "=" + value);
                Logs.AddLogsCollected($"Config added: " + parameter + "=" + value);
            }
        }

        public static string HtmlFilePath { get; set; }

        public static string XlsFilePath { get; set; }

        public static bool FiveDayMode { get; set; }

        public static string BirthdayColumnNumber { get; set; }
        public static string EmployeeNameColumnNumber { get; set; }             

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

        public static void LoadConfig()
        {
            ParametersList = new List<string>(File.ReadAllLines(ConfigsPath));
            foreach (var item in ParametersList)
            {
                string parameter = item.Substring(0, item.IndexOf('='));
                string value = item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1);
                switch (parameter)
                {
                    case "senderEmail":
                        SenderEmail = value;
                        break;
                    case "senderUsername":
                        SenderUsername = value;
                        break;
                    case "senderPassword":
                        SenderPassword = value;
                        break;
                    case "senderName":
                        SenderName = value;
                        break;
                    case "emailRecievers":
                        EmailRecievers = new List<string>(value.Split(','));
                        break;
                    case "messageSubject":
                        MessageSubject = value;
                        break;
                    case "htmlPath":
                        HtmlFilePath = value;
                        break;
                    case "xlsPath":
                        XlsFilePath = value;
                        break;
                    case "birthdayColumnNumber":
                        BirthdayColumnNumber = value;
                        break;
                    case "employeeNameColumnNumber":
                        EmployeeNameColumnNumber = value;
                        break;
                    case "serverAddress":
                        ServerAddress = value;
                        break;
                    case "serverPort":
                        ServerPort = Convert.ToInt32(value);
                        break;
                    case "fiveDaysMode":
                        if (value.ToLower() == "yes" ||
                        value.ToLower() == "y" ||
                        value.ToLower() == "true")
                        {
                            FiveDayMode = true;
                        }
                        else
                        {
                            FiveDayMode = false;
                        }
                        break;
                    case "logRecievers":
                        LogsRecievers = new List<string>(value.Split(','));
                        break;
                    default:
                        break;
                }
            }
        }

        public static void EditConfig(string parameter, string value)
        {
            string fileType;
            switch (parameter)
            {
                case "birthdayColumnNumber":
                case "employeeNameColumnNumber":
                    if (!int.TryParse(value, out _))
                    {
                        value = "";
                    }
                    break;
                case "serverPort":
                    if (string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(value))
                    {
                        value = "25";
                    }
                    else if (!int.TryParse(value, out _))
                    {
                        value = "";
                    }
                    break;
                case "htmlPath":
                    fileType = value.Substring(value.LastIndexOf('.') + 1, value.Length - value.LastIndexOf('.') - 1);
                    if (File.Exists(value) && fileType.ToLower().Equals("html"))
                    {
                        if (!File.ReadAllText(value).Contains("%LIST_OF_EMPLOYEES%"))
                        {
                            value = "";
                        }
                    }
                    else
                    {
                        value = "";
                    }
                    break;
                case "xlsPath":
                    fileType = value.Substring(value.LastIndexOf('.') + 1, value.Length - value.LastIndexOf('.') - 1);
                    if (!File.Exists(value) || !fileType.ToLower().Equals("xls"))
                    {
                        value = "";
                    }
                    break;
                case "senderPassword":
                    value = Encryptor.EncryptString("b14ca5898a4e4133bbce2mbd02082020", value);
                    break;
                default:
                    break;
            }
            ChangeParameter(parameter, value);
        }

        public static void SaveConfig()
        {
            try
            {
                File.WriteAllText(ConfigsPath, string.Empty);
                File.WriteAllLines(ConfigsPath, ParametersList);
                Logs.AddLogsCollected($"Config save: SUCCESS.");
            }
            catch
            {
                Logs.AddLogsCollected($"Config save: FAILURE.");
            }
        }
    }
}