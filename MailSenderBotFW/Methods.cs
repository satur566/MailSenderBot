using ExcelLibrary.SpreadSheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Security.Cryptography;
using System.Text;

namespace MailSender
{
    static class Methods
    {
        public static void SendMail()
        {
            ReadXlsFile(Configs.XlsFilePath, Configs.FiveDayMode, Configs.BirthdayColumnNumber, Configs.EmployeeNameColumnNumber);
            ReadHtmlFile(Configs.HtmlFilePath, Employees.CongratulationsList);
            if (Employees.WhosBirthdayIs.Count.Equals(0) || Configs.FiveDayMode && (DateTime.Now.DayOfWeek == DayOfWeek.Sunday || DateTime.Now.DayOfWeek == DayOfWeek.Saturday))
            {
                Configs.AddLogsCollected($"Sending message: CANCELLED.");
                if (Configs.FiveDayMode && (DateTime.Now.DayOfWeek == DayOfWeek.Sunday || DateTime.Now.DayOfWeek == DayOfWeek.Saturday))
                {
                    Configs.AddLogsCollected($"Reason: today is a day off.");
                }
                if (Employees.WhosBirthdayIs.Count.Equals(0))
                {
                    Configs.AddLogsCollected($"Reason: employees don't have a birthday today.");
                } else if (Configs.FiveDayMode && (DateTime.Now.DayOfWeek == DayOfWeek.Sunday || DateTime.Now.DayOfWeek == DayOfWeek.Saturday))
                {
                    Configs.AddLogsCollected($"On Monday {Employees.WhosBirthdayIs.Count} employees will be congratulated.");
                }                
            }
            else
            {
                if (Configs.EmailRecievers.Count.Equals(0))
                {
                    Configs.AddLogsCollected($"Sending message: CANCELLED.");
                    Configs.AddLogsCollected($"Reason: recievers count equals 0.");
                }
                else
                {
                    foreach (var reciever in Configs.EmailRecievers)
                    {
                        SendMessage(reciever, Configs.MessageSubject, Configs.MessageText);
                    }
                }
            }
            SendLogs();
        }

        private static void SendMessage(string reciever, string subject, string message)
        {
            try
            {
                MailAddress Sender = new MailAddress(Configs.SenderEmail, Configs.SenderName);
                MailAddress Reciever = new MailAddress(reciever);
                MailMessage Message = new MailMessage(Sender, Reciever)
                {
                    Subject = subject,
                    Body = message,
                    IsBodyHtml = true
                };
                List<LinkedResource> images = new List<LinkedResource>();
                string[] htmlArray = message.Split(new string[] { System.Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                string htmlFolderPath = Configs.HtmlFilePath.Substring(0, Configs.HtmlFilePath.LastIndexOf('\\') + 1);
                int imageCounter = 0;
                foreach (var line in htmlArray)
                {
                    if (line.Contains("src="))
                    {
                        string srcLine = line.Substring(line.IndexOf('\"') + 1, line.Substring(line.IndexOf('\"') + 1).IndexOf('\"'));
                        string imagePath = htmlFolderPath + srcLine.Replace('/', '\\');
                        images.Add(new LinkedResource(imagePath, "image/gif"));
                        message = message.Replace(srcLine, "cid:" + images[imageCounter++].ContentId);
                    }
                }
                var htmlView = AlternateView.CreateAlternateViewFromString(message, Encoding.UTF8, MediaTypeNames.Text.Html);
                images.ForEach(htmlView.LinkedResources.Add);
                Message.AlternateViews.Add(htmlView);  //I FORGOT SOMETHING!!!!
                SmtpClient Client = new SmtpClient(Configs.ServerAddress, Configs.ServerPort)
                {
                    Credentials = new NetworkCredential(Configs.SenderUsername, DecryptString("b14ca5898a4e4133bbce2mbd02082020", Configs.SenderPassword)),
                    EnableSsl = false
                };
                Client.Send(Message);
                if (!subject.Contains("log from"))
                {
                    Configs.AddLogsCollected($"Sending message to {reciever}: SUCCESS.");
                    Configs.AddLogsCollected(LogConclusionMaker(reciever));
                }
            }
            catch
            {
                if (!subject.Contains("log from"))
                {
                    Configs.AddLogsCollected($"Sending message to {reciever}: FAILURE.");
                } else
                {
                    Configs.AddLogsCollected($"Sending log to {reciever}: FAILURE.");
                }
            }
        }

        private static void SendLogs()
        {
            foreach (var reciever in Configs.LogsRecievers)
            {
                string ifSuccess = "";
                try
                {
                    SendMessage(reciever, $"log from {DateTime.Now}", Configs.LogsCollected);
                    ifSuccess = "SUCCESS";
                }
                catch
                {
                    ifSuccess = "FAILURE";
                }
                finally
                {
                    Configs.AddLogsCollected($"Sending logs: {ifSuccess}.");
                }
            }
        }

        private static string LogConclusionMaker(string reciever)
        {
            string employees = "";
            foreach (var item in Employees.WhosBirthdayIs)
            {
                employees = String.Concat(employees, "\t" + item.Trim() + "\n");

            }
            string result = $"\nConclusion:" +
                $"\nSender mail: {Configs.SenderEmail}" +
                $"\nSender name: {Configs.SenderName}" +
                $"\nReciever e-mail:{reciever}" +
                $"\nBirthday givers:\n{employees}";
            return result;
        }

        private static void ReadXlsFile(string path, bool fiveDaysMode, string birthdayColumn, string employeeColumn)
        {
            bool isNextMonthLoaded = false;
            bool isTempCopied = false;
            try
            {
                string fileName = path.Substring(path.LastIndexOf('\\'));
                File.Copy(path, Configs.TempPath + fileName, true);
                path = Configs.TempPath + fileName;
                isTempCopied = true;
            }
            catch 
            {
                Configs.AddLogsCollected("Unable to open temporary copy of existing .xls file.");
            }
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
                            if (!isNextMonthLoaded && Convert.ToDateTime(BirthdaySheet.Cells[i, bdColumn].ToString()).Date.Month.Equals(DateTime.Now.Month + 1))
                            {
                                isNextMonthLoaded = true;
                                break;
                            }
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
                            Employees.AddBirthdaygiver(BirthdaySheet.Cells[i, emColumn].ToString());
                        }
                        if (DateTime.Now.DayOfWeek == DayOfWeek.Monday && fiveDaysMode)
                        {
                            try
                            {
                                if (Convert.ToDateTime(BirthdaySheet.Cells[i, bdColumn].ToString()).Date.Equals(DateTime.Now.AddDays(-1).Date))
                                {
                                    Employees.AddBirthdaygiver(BirthdaySheet.Cells[i, emColumn].ToString());
                                }
                            }
                            catch { }
                            try
                            {
                                if (Convert.ToDateTime(BirthdaySheet.Cells[i, bdColumn].ToString()).Date.Equals(DateTime.Now.AddDays(-2).Date))
                                {
                                    Employees.AddBirthdaygiver(BirthdaySheet.Cells[i, emColumn].ToString());
                                }
                            }
                            catch { }
                        }
                    }                    
                    catch
                    {
                        continue;
                    }
                    try
                    {                        
                        
                    }
                    catch
                    {
                        continue;
                    }
                }
                Configs.AddLogsCollected($"Reading xls: SUCCESS.");
                if (DateTime.Now.Day > 25 && !isNextMonthLoaded)
                {
                    Configs.AddLogsCollected("Xls file does not contain list of employees for a next month.");
                }
                if (isTempCopied)
                {
                    try
                    {
                        File.Delete(path);
                    }
                    catch
                    {
                        Configs.AddLogsCollected("Unable to delete temporary .xls file.");
                    }
                }
            }
            catch
            {
                Configs.AddLogsCollected($"Reading xls: FAILURE.");
            }
        }

        private static void ReadHtmlFile(string path, string employees)
        {
            if (File.ReadAllText(path).Contains("%LIST_OF_EMPLOYEES%"))
            {
                Configs.MessageText = File.ReadAllText(path).Replace("%LIST_OF_EMPLOYEES%", employees);
                Configs.AddLogsCollected($"Reading html: SUCCESS.");
            }
            else
            {
                Configs.AddLogsCollected($"Reading html: FAILURE.");
                Configs.AddLogsCollected($"Reason: list of employees can't be inserted.");
            }
        }

        public static void LoadConfig()
        {
            Configs.ParametersList = new List<string>(File.ReadAllLines(Configs.ConfigsPath));
            foreach (var item in Configs.ParametersList)
            {
                string parameter = item.Substring(0, item.IndexOf('='));
                string value = item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1);
                switch (parameter)
                {
                    case "senderEmail":
                        Configs.SenderEmail = value;
                        break;
                    case "senderUsername":
                        Configs.SenderUsername = value;
                        break;
                    case "senderPassword":
                        Configs.SenderPassword = value;
                        break;
                    case "senderName":
                        Configs.SenderName = value;
                        break;
                    case "emailRecievers":
                        Configs.EmailRecievers = new List<string>(value.Split(','));
                        break;
                    case "messageSubject":
                        Configs.MessageSubject = value;
                        break;
                    case "htmlPath":
                        Configs.HtmlFilePath =value;
                        break;
                    case "xlsPath":
                        Configs.XlsFilePath = value;
                        break;
                    case "birthdayColumnNumber":
                        Configs.BirthdayColumnNumber = value;
                        break;
                    case "employeeNameColumnNumber":
                        Configs.EmployeeNameColumnNumber =value;
                        break;
                    case "serverAddress":
                        Configs.ServerAddress = value;
                        break;
                    case "serverPort":
                        Configs.ServerPort = Convert.ToInt32(value);
                        break;
                    case "fiveDaysMode":
                        Configs.FiveDayMode = Boolean.TryParse(value, out _);
                        break;
                    case "logRecievers":
                        Configs.LogsRecievers = new List<string>(value.Split(','));
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
                    if (!Int32.TryParse(value, out _))
                    {
                        value = "";
                    }
                    break;
                case "serverPort":
                    if (string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(value))
                    {
                        value = "25";
                    }
                    else if (!Int32.TryParse(value, out _))
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
                    value = EncryptString("b14ca5898a4e4133bbce2mbd02082020", value);
                    break;
                case "fiveDaysMode":
                    if (value.ToLower() == "yes" || value.ToLower() == "y")
                    {
                        value = "True";
                    }
                    else
                    {
                        value = "False";
                    }
                    break;
                default:
                    break;
            }
            Configs.ChangeParameter(parameter, value);            
        }

        public static void SaveConfig()
        {
            try
            {
                File.WriteAllText(Configs.ConfigsPath, string.Empty);
                File.WriteAllLines(Configs.ConfigsPath, Configs.ParametersList);
                Configs.AddLogsCollected($"Config save: SUCCESS.");
            }
            catch
            {
                Configs.AddLogsCollected($"Config save: FAILURE.");
            }
        }

        private static string EncryptString(string key, string plainText)
        {
            byte[] iv = new byte[16];
            byte[] array;

            using (Aes aes = Aes.Create())
            {
                aes.Key = Encoding.UTF8.GetBytes(key);
                aes.IV = iv;

                ICryptoTransform encryptor = aes.CreateEncryptor(aes.Key, aes.IV);

                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (CryptoStream cryptoStream = new CryptoStream((Stream)memoryStream, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter streamWriter = new StreamWriter((Stream)cryptoStream))
                        {
                            streamWriter.Write(plainText);
                        }

                        array = memoryStream.ToArray();
                    }
                }
            }

            return Convert.ToBase64String(array);
        }

        private static string DecryptString(string key, string cipherText)
        {
            byte[] iv = new byte[16];
            byte[] buffer = Convert.FromBase64String(cipherText);

            using (Aes aes = Aes.Create())
            {
                aes.Key = Encoding.UTF8.GetBytes(key);
                aes.IV = iv;
                ICryptoTransform decryptor = aes.CreateDecryptor(aes.Key, aes.IV);

                using (MemoryStream memoryStream = new MemoryStream(buffer))
                {
                    using (CryptoStream cryptoStream = new CryptoStream((Stream)memoryStream, decryptor, CryptoStreamMode.Read))
                    {
                        using (StreamReader streamReader = new StreamReader((Stream)cryptoStream))
                        {
                            return streamReader.ReadToEnd();
                        }
                    }
                }
            }
        }
    }
}