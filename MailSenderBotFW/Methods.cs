using System;
using System.Collections.Generic;
using ExcelLibrary.SpreadSheet;
using System.Net.Mail;
using System.IO;
using System.Net;

namespace MailSender
{
    static class Methods
    {
        public static void SendMail()
        {
            if (Employees.GetWhosBirthdayIs().Count.Equals(0) || Configs.GetFiveDayMode() && (DateTime.Now.DayOfWeek == DayOfWeek.Sunday || DateTime.Now.DayOfWeek == DayOfWeek.Saturday))
            {
                Configs.AddLogsCollected($"Sending message: CANCELLED.");
                if (Employees.GetWhosBirthdayIs().Count.Equals(0))
                {
                    Configs.AddLogsCollected($"Reason: employees don't have a birthday today.");                    
                }
                if (Configs.GetFiveDayMode() && (DateTime.Now.DayOfWeek == DayOfWeek.Sunday || DateTime.Now.DayOfWeek == DayOfWeek.Saturday))
                {
                    Configs.AddLogsCollected($"Reason: today is a day off.");
                }
            }
            else
            {                
                SendMessage(Configs.GetRecieverEmail(), Configs.GetMessageSubject(), Configs.GetMessageText());
                Configs.AddLogsCollected(LogConclusionMaker()); 
            }
            SendLogs();
        }

        private static void SendLogs()
        {
            foreach (var reciever in Configs.GetLogRecievers())
            {
                try
                {
                    SendMessage(reciever, $"log from {DateTime.Now}", Configs.GetLogsCollected());
                    Configs.AddLogsCollected($"Sending logs: SUCCESS.");
                }
                catch
                {
                    Configs.AddLogsCollected($"Sending logs: FAILURE.");
                }
            }
        }

        private static string LogConclusionMaker()
        {
            string employees = "";
            foreach (var item in Employees.GetWhosBirthdayIs())
            {
                employees = employees + item.Trim() + "\n";

            }
            string result = $"\nConclusion:" +
                $"\nSender mail: {Configs.GetSenderEmail()}" +
                $"\nSender name: {Configs.GetSenderName()}" +
                $"\nReciever e-mail: {Configs.GetRecieverEmail()}" +
                $"\nBirthday givers:\n{employees}";
            return result;
        }

        public static void LoadConfig() //TODO: Is all config loaded?
        {
            Configs.SetConfigurations(new List<string>(File.ReadAllLines(Configs.GetConfigPath())));
            foreach (var item in Configs.GetConfigurations())
            {
                string parameter = item.Substring(0, item.IndexOf('='));
                string value = item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1);
                switch (parameter)
                {
                    case "senderEmail":
                        Configs.SetSenderEmail(value);
                        break;
                    case "senderPassword":
                        Configs.SetSenderPassword(value);
                        break;
                    case "senderName":
                        Configs.SetSenderName(value);
                        break;
                    case "recieverEmail":
                        Configs.SetRecieverEmail(value);
                        break;
                    case "messageSubject":
                        Configs.SetMessageSubject(value);
                        break;
                    case "htmlPath":
                        Configs.SetHtmlPath(value);
                        break;
                    case "xlsPath":
                        Configs.SetXlsPath(value);
                        break;
                    case "birthdayColumnNumber":
                        Configs.SetBirthdayColumnNumber(value);
                        break;
                    case "employeeNameColumnNumber":
                        Configs.SetEmployeeNameColumnNumber(value);
                        break;
                    case "serverAddress":
                        Configs.SetServerAddress(value);
                        break;
                    case "serverPort":
                        Configs.SetServerPort(value);
                        break;
                    case "fiveDaysMode":
                        Configs.SetFiveDayMode(Boolean.TryParse(value, out _));
                        break;
                    case "logRecievers":
                        string[] recievers = value.Split(',');
                        foreach (var reciever in recievers)
                        {
                            Configs.SetLogRecievers(reciever.Trim());
                        }
                        break;
                    default:
                        break;
                }
            }
            Configs.SetMessageText(File.ReadAllText(Configs.GetHtmlPath()));
            ReadXlsFile(Configs.GetXlsPath(), Configs.GetFiveDayMode(), Configs.GetBirthdayColumnNumber(), Configs.GetEmployeeNameColumnNumber());
            ReadHtmlFile(Configs.GetHtmlPath(), Employees.GetCongratulationsString());
            Configs.SetReadConfigSuccess(true);
        }

        public static void ReadHtmlFile(string path, string employees)
        {
            if (File.ReadAllText(path).Contains("%LIST_OF_EMPLOYEES%"))
            {
                Configs.SetMessageText(File.ReadAllText(path).Replace("%LIST_OF_EMPLOYEES%", employees));
                Configs.AddLogsCollected($"Reading html: SUCCESS.");
            }
            else
            {
                Configs.AddLogsCollected($"Reading html: FAILURE.");
                Configs.AddLogsCollected($"Reason: list of employees can't be inserted.");
            }
        }

        private static void SendMessage(string reciever, string subject, string message)
        {
            try
            {
                MailAddress Sender = new MailAddress(Configs.GetSenderEmail(), Configs.GetSenderName());
                MailAddress Reciever = new MailAddress(reciever);
                MailMessage Message = new MailMessage(Sender, Reciever)
                {
                    Subject = subject,
                    Body = message,
                    IsBodyHtml = true
                };
                SmtpClient Client = new SmtpClient(Configs.GetServerAddress(), Convert.ToInt32(Configs.GetServerPort()))
                {
                    Credentials = new NetworkCredential(Configs.GetSenderUsername(), Configs.GetSenderPassword()),
                    EnableSsl = false
                };
                Client.Send(Message);
                if (!subject.Contains("log from"))
                {
                    Configs.AddLogsCollected($"Sending message: SUCCESS.");
                }
            }
            catch
            {
                Configs.AddLogsCollected($"Sending message: FAILURE.");
            }
        }
        public static void EditConfig(string type, string parameter, string value) //TODO: return value.
        {
            string fileType = value.Substring(value.LastIndexOf('.') + 1, value.Length - value.LastIndexOf('.') - 1); //TODO: under html/xls
            switch (type) //TODO: switch parameter multicase for only text values etc.
            {
                case "digit":
                    if (!Int32.TryParse(value, out _))
                    {
                        value = "";
                    }                        
                    break;
                case "port":
                    if (string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(value))
                    {
                        value = "25";
                    } 
                    else if (!Int32.TryParse(value, out _))
                    {
                        value = "";
                    }                    
                    break;
                case "html":                    
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
                case "xls":                    
                    if (!File.Exists(value) || !fileType.ToLower().Equals("xls"))
                    {
                        value = "";
                    }
                    break;
                case "password": //TODO: encode.

                    break;
                case "y/n": //TODO this!
                    if (value.ToLower() == "yes" || value.ToLower() == "y")
                    {
                        value = "True";
                    } else
                    {
                        value = "False";
                    }
                    break;
                default:
                    break;
            }
            Configs.SetConfigurations(parameter + "=" + value);
            Configs.AddLogsCollected($"Config: " + parameter + "=" + value);
        }

        public static void SaveConfig()
        {                        
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
            ReadXlsFile(Configs.GetXlsPath(), Configs.GetFiveDayMode(), Configs.GetBirthdayColumnNumber(), Configs.GetEmployeeNameColumnNumber()); //TODO: only saveConfig methods.
            ReadHtmlFile(Configs.GetHtmlPath(), Employees.GetCongratulationsString());
            LoadConfig();
            Configs.SetReadConfigSuccess(false);            
        }        

        public static void ReadXlsFile(string path, bool fiveDaysMode, string birthdayColumn, string employeeColumn) //TODO: read through locked file. 
        {
            try
            {
                int bdColumn = Convert.ToInt32(birthdayColumn);
                int emColumn = Convert.ToInt32(employeeColumn);
                bdColumn--;
                emColumn--;
                Workbook BirthdayBook = Workbook.Load(path);
                Worksheet BirthdaySheet = BirthdayBook.Worksheets[0];
                for (int i = 0; i < BirthdaySheet.Cells.LastRowIndex; i++)
                {
                    try
                    {
                        if (Convert.ToDateTime(BirthdaySheet.Cells[i, bdColumn].ToString()).Date > DateTime.Today)
                        {
                            continue;
                        }
                    }
                    catch
                    {
                        continue;
                    }
                    try
                    {
                        if (Convert.ToDateTime(BirthdaySheet.Cells[i, bdColumn].ToString()).Date.Equals(DateTime.Now.Date))
                        {
                            Employees.SetWhosBirthdayIs(BirthdaySheet.Cells[i, emColumn].ToString());
                        }
                        if (DateTime.Now.DayOfWeek == DayOfWeek.Monday && fiveDaysMode)
                        {
                            try
                            {
                                if (Convert.ToDateTime(BirthdaySheet.Cells[i, bdColumn].ToString()).Date.Equals(DateTime.Now.AddDays(-1).Date))
                                {
                                    Employees.SetWhosBirthdayIs(BirthdaySheet.Cells[i, emColumn].ToString());
                                }
                            }
                            catch { }
                            try
                            {
                                if (Convert.ToDateTime(BirthdaySheet.Cells[i, bdColumn].ToString()).Date.Equals(DateTime.Now.AddDays(-2).Date))
                                {
                                    Employees.SetWhosBirthdayIs(BirthdaySheet.Cells[i, emColumn].ToString());
                                }
                            }
                            catch { }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
                Configs.AddLogsCollected($"Reading xls: SUCCESS.");
            } 
            catch
            {
                Configs.AddLogsCollected($"Reading xls: FAILURE.");
            }
        }
    }
}
