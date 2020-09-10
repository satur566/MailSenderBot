using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace MailSender
{
    static class Configs //TODO: add responsibleUsers and send them message about fulfilling of xls list.
    {
        private static string senderEmail;
        private static string senderPassword;
        private static string serverPort;
        private static string htmlFilePath;
        private static string xlsFilePath;
        private static string birthdayColumnNumber;
        private static string employeeNameColumnNumber;
        private static readonly List<string> htmlFilesList = new List<string>();
        private static readonly List<string> emailRecievers = new List<string>();
        private static List<string> parametersList = new List<string>();
        private static readonly List<string> logRecievers = new List<string>();            

        //TODO: make replaceable string %CONGRATULATION_TEXT% similar as %LIST_OF_USERS% and reads that text from txt file.
        //TODO: make property of congratulationText path
         
        //TODO: switch html samples like: every day, every week, every month, every time of the year.

        public static string SenderEmail
        {
            get
            {
                return senderEmail;
            }
            set //TODO: add this to recievers and log recievers without throw
            {
                if (string.IsNullOrEmpty(value) 
                    && string.IsNullOrWhiteSpace(value))
                {
                    string exceptionMessage = "Wrong parameter value. Email can't be empty.";
                    Logs.AddLogsCollected(exceptionMessage);
                    throw new Exception(exceptionMessage);
                }
                if (!value.Contains("@"))
                {
                    string exceptionMessage = "Wrong parameter value. Email should have an @ symbol.";
                    Logs.AddLogsCollected(exceptionMessage);
                    throw new Exception(exceptionMessage);
                }
                senderEmail = value;
            }
        }

        public static string SenderName { set; get; }

        public static string SenderUsername { set; get; }

        public static string SenderPassword {
            get
            {
                return Encryptor.DecryptString("b14ca5898a4e4133bbce2mbd02082020", senderPassword);
            }
            set
            {
                value = Encryptor.EncryptString("b14ca5898a4e4133bbce2mbd02082020", value);
                senderPassword = value;
            }
        }

        public static List<string> EmailRecievers //TODO: throw if .count.equals 0
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

        public static string ServerAddress { set; get; } //TODO: throw if empty

        public static string ServerPort {
            get
            {
                return serverPort;
            }
            set
            {
                if (string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(value))
                {
                    value = "25";
                }
                else if (!int.TryParse(value, out _))
                {
                    string exceptionMessage = "Wrong parameter value. Port number should be a digit!";
                    Logs.AddLogsCollected(exceptionMessage);
                    throw new Exception(exceptionMessage);
                }
                serverPort = value;
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
                SortConfiguration(ref tempList, "htmlPath");
                SortConfiguration(ref tempList, "htmlFolderPath");
                SortConfiguration(ref tempList, "htmlSwitchMode");
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
                parametersList.Clear();
                foreach (var item in value)
                {
                    string parameter = item.Substring(0, item.IndexOf('='));
                    string parameterValue = item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1);
                    switch (parameter)
                    {
                        case "senderEmail":
                            SenderEmail = parameterValue;
                            break;
                        case "senderUsername":
                            SenderUsername = parameterValue;
                            break;
                        case "senderPassword":
                            SenderPassword = parameterValue;
                            break;
                        case "senderName":
                            SenderName = parameterValue;
                            break;
                        case "emailRecievers":
                            EmailRecievers = new List<string>(parameterValue.Split(','));
                            break;
                        case "messageSubject":
                            MessageSubject = parameterValue;
                            break;
                        case "htmlPath":
                            HtmlFilePath = parameterValue;
                            break;
                        case "htmlFolderPath":
                            HtmlFolderPath = parameterValue;
                            break;
                        case "htmlSwitchMode":
                            HtmlSwitchMode = parameterValue;
                            break;
                        case "xlsPath":
                            XlsFilePath = parameterValue;
                            break;
                        case "birthdayColumnNumber":
                            BirthdayColumnNumber = parameterValue;
                            break;
                        case "employeeNameColumnNumber":
                            EmployeeNameColumnNumber = parameterValue;
                            break;
                        case "serverAddress":
                            ServerAddress = parameterValue;
                            break;
                        case "serverPort":
                            ServerPort = parameterValue;
                            break;
                        case "fiveDaysMode":
                            if (parameterValue.ToLower() == "yes" ||
                            parameterValue.ToLower() == "y" ||
                            parameterValue.ToLower() == "true")
                            {
                                FiveDayMode = true;
                            }
                            else
                            {
                                FiveDayMode = false;
                            }
                            break;
                        case "logRecievers":
                            LogsRecievers = new List<string>(parameterValue.Split(','));
                            break;
                        default:
                            break;
                    }
                }
                parametersList = value;
            }
        }

        public static string HtmlFilePath { 
            get
            {
                return htmlFilePath;
            }
            set
            {
                string fileType = value.Substring(value.LastIndexOf('.') + 1, value.Length - value.LastIndexOf('.') - 1);
                if (!File.Exists(value))
                {
                    string exceptionMessage = "Wrong parameter value. Current path does not lead to existing file.";
                    Logs.AddLogsCollected(exceptionMessage);
                    throw new Exception(exceptionMessage);
                } 
                else if (!fileType.ToLower().Equals("html"))
                {
                    string exceptionMessage = "Wrong parameter value. File extension is not html.";
                    Logs.AddLogsCollected(exceptionMessage);
                    throw new Exception(exceptionMessage);
                } 
                else if ((!File.ReadAllText(value).Contains("%LIST_OF_EMPLOYEES%"))) {
                    string exceptionMessage = "Wrong parameter value. Html-file does not contain string: %LIST_OF_EMPLOYEES%";
                    Logs.AddLogsCollected(exceptionMessage);
                    throw new Exception(exceptionMessage);
                } 
                else
                {
                    htmlFilePath = value;
                }                
            }
        }

        public static string HtmlFolderPath { get; set; } //TODO: throw if folder !exist or does not contain html files

        public static string HtmlSwitchMode { get; set; }

        public static List<string> HtmlFilesList
        {
            get
            {
                htmlFilesList.Sort();
                return htmlFilesList;
            }
            set
            {
                htmlFilesList.Clear();
                foreach (string htmlFile in value)
                {
                    htmlFilesList.Add(htmlFile);
                }
            }
        }

        public static string XlsFilePath { 
            get
            {
                return xlsFilePath;
            }
            set
            {
                string fileType = value.Substring(value.LastIndexOf('.') + 1, value.Length - value.LastIndexOf('.') - 1);
                if (!fileType.ToLower().Equals("xls"))
                {
                    string exceptionMessage = "Wrong parameter value. File extension is not xls.";
                    Logs.AddLogsCollected(exceptionMessage);
                    throw new Exception(exceptionMessage);
                }
                else if (!File.Exists(value))
                {
                    string exceptionMessage = "Wrong parameter value. Current path does not lead to existing file.";
                    Logs.AddLogsCollected(exceptionMessage);
                    throw new Exception(exceptionMessage);
                } else
                {
                    xlsFilePath = value;
                }
            }
        }

        public static bool FiveDayMode { get; set; }

        public static string BirthdayColumnNumber { 
            get
            {
                return birthdayColumnNumber;
            }
            set 
            {
                if (!int.TryParse(value, out _))
                {
                    string exceptionMessage = "Wrong parameter value. Column number should be a digit!";
                    Logs.AddLogsCollected(exceptionMessage);
                    throw new Exception(exceptionMessage);
                }
                birthdayColumnNumber = value;
            } 
        }
        public static string EmployeeNameColumnNumber {
            get
            {
                return employeeNameColumnNumber;
            } 
            set
            {
                if (!int.TryParse(value, out _))
                {
                    string exceptionMessage = "Wrong parameter value. Column number should be a digit!";
                    Logs.AddLogsCollected(exceptionMessage);
                    throw new Exception(exceptionMessage);
                }
                employeeNameColumnNumber = value;
            }
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

        public static void EditConfig(string parameter, string value)
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

        public static string RandomChangeHtmlFile(string currentHtmlFilePath)
        {
            HtmlFilesList = FileWorks.CollectHtmlFiles(HtmlFolderPath);
            Random random = new Random();
            int selectedIndex;
            while (true)
            {
                selectedIndex = random.Next(0, HtmlFilesList.Count);
                if (!HtmlFilesList[selectedIndex].Equals(currentHtmlFilePath))
                {
                    break;
                }
            }
            return HtmlFilesList[selectedIndex];
        }

        public static string AscendingChangeHtmlFile(string currentHtmlFilePath)
        {
            HtmlFilesList = FileWorks.CollectHtmlFiles(HtmlFolderPath);
            int selectedIndex = HtmlFilesList.IndexOf(currentHtmlFilePath) + 1;
            if (selectedIndex.Equals(HtmlFilesList.Count))
            {
                selectedIndex = 0;
            }
            return HtmlFilesList[selectedIndex];
        }
    }
}