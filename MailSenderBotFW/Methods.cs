using System;
using System.Collections.Generic;
using System.Linq;
using ExcelLibrary.SpreadSheet;
using System.Net.Mail;
using System.IO;
using System.Net;

namespace MailSender
{
    static class Methods //TODO: remove all console methods.
    {
        public static void SendMail(string[] args)
        {
            if (String.IsNullOrEmpty(ShowBirthdayGivers(false, false)) || Configs.GetFiveDayMode() && (DateTime.Now.DayOfWeek == DayOfWeek.Sunday || DateTime.Now.DayOfWeek == DayOfWeek.Saturday))
            {
                Configs.AddLogsCollected($"Sending message: CANCELLED.");
                if (String.IsNullOrEmpty(ShowBirthdayGivers(false, false)))
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
                Configs.SetMessageText(File.ReadAllText(Configs.GetHtmlPath())); //Throw somewhere
                Console.Write($"\nConclusion:" +
                    $"\nSender mail: {Configs.GetSenderEmail()}" +
                    $"\nSender name: {Configs.GetSenderName()}" +
                    $"\nReciever e-mail: {Configs.GetRecieverEmail()}" +
                    $"\n" +
                    $"\n{ShowBirthdayGivers(false, false)}\n");

                if (Configs.GetReadConfigSuccess() && args.Contains<string>("-silent"))
                {
                    SendMessage(Configs.GetRecieverEmail(), Configs.GetMessageSubject(), Configs.GetMessageText(), args, true);
                    Configs.AddLogsCollected($"Sending message mode: silent");                    
                }
                else
                {
                    Console.Write("\nIs everything fine? (Y/N)");
                    switch (Console.ReadLine().ToLower())
                    {
                        case "y":
                            SendMessage(Configs.GetRecieverEmail(), Configs.GetMessageSubject(), Configs.GetMessageText(), args, true);
                            break;
                        default:
                            Configs.AddLogsCollected($"Sending message: CANCELLED.");
                            Configs.AddLogsCollected($"Reason: user cancel.");
                            break;
                    }
                }
            }
            SendLogs(args);
        }

        private static void SendLogs(string[] args)
        {
            foreach (var reciever in Configs.GetLogRecievers())
            {
                try
                {
                    SendMessage(reciever, $"log from {DateTime.Now}", Configs.GetLogsCollected(), args, false);
                    Configs.AddLogsCollected($"Sending logs: SUCCESS.");
                }
                catch
                {
                    Configs.AddLogsCollected($"Sending logs: FAILURE.");
                }
            }
        }

        private static string ShowBirthdayGivers(bool inLogs, bool onEmail) //TODO: send conclusion method.
        {
            string result = "";
            foreach (var item in Employees.GetWhosBirthdayIs())
            {
                if (inLogs)
                {
                    if (onEmail)
                    {
                        result = result + item + "\n\t\t\t\t\t\t<br>";
                    }
                    else
                    {
                        result = result + item + "\n\t\t\t\t\t\t";
                    }

                }
                else
                {
                    result = result + item + "\n";
                }
            }
            return result.Trim();
        }

        public static void ReadConfig()
        {
            Configs.SetConfigurations(new List<string>(File.ReadAllLines(Configs.GetConfigPath())));
            foreach (var item in Configs.GetConfigurations())
            {
                switch (item.Substring(0, item.IndexOf('=')))
                {
                    case "senderEmail":
                        Configs.SetSenderEmail(item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1));
                        break;
                    case "senderPassword":
                        Configs.SetSenderPassword(item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1));
                        break;
                    case "senderName":
                        Configs.SetSenderName(item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1));
                        break;
                    case "senderUsername":
                        Configs.SetSenderUsername(item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1));
                        break;
                    case "recieverEmail":
                        Configs.SetRecieverEmail(item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1));
                        break;
                    case "messageSubject":
                        Configs.SetMessageSubject(item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1));
                        break;
                    case "htmlPath":
                        Configs.SetHtmlPath(item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1));
                        break;
                    case "xlsPath":
                        Configs.SetXlsPath(item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1));
                        break;
                    case "birthdayColumnNumber":
                        Configs.SetBirthdayColumnNumber(item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1));
                        break;
                    case "employeeNameColumnNumber":
                        Configs.SetEmployeeNameColumnNumber(item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1));
                        break;
                    case "serverAddress":
                        Configs.SetServerAddress(item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1));
                        break;
                    case "serverPort":
                        Configs.SetServerPort(item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1));
                        break;
                    case "fiveDaysMode":
                        Configs.SetFiveDayMode(Boolean.TryParse(item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1), out _));
                        break;
                    case "logRecievers":
                        string[] recievers = item.Substring(item.IndexOf('=') + 1, item.Length - item.IndexOf('=') - 1).Split(',');
                        foreach (var reciever in recievers)
                        {
                            Configs.SetLogRecievers(reciever.Trim());
                        }
                        break;
                    default:
                        break;
                }
            }          
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
                Console.WriteLine("HTML file should have %LIST_OF_EMPLOYEES% string. Current file has no such string. Continue? (y/n)");
                switch (Console.ReadLine().ToLower())
                {
                    case "y":
                        Configs.AddLogsCollected($"Reading html: SUCCESS.Html file dous not contain list of employees.");
                        break;
                    default:
                        Configs.AddLogsCollected($"Reading html: FAILURE.");
                        Environment.Exit(0);
                        break;
                }
            }
        }

        private static void SendMessage(string reciever, string subject, string message, string[] args, bool enableLog)
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
                if (enableLog)
                {
                    Configs.AddLogsCollected($"Sending message: SUCCESS."); //TODO: send conclusion.
                    Configs.AddLogsCollected($"Conclusion: <br>" +
                        $"\n\t\t\t\t\t\tSender mail: {Configs.GetSenderEmail()}<br>" +
                        $"\n\t\t\t\t\t\tSender name: {Configs.GetSenderName()}<br>" +
                        $"\n\t\t\t\t\t\tReciever e-mail: {Configs.GetRecieverEmail()}<br>" +
                        $"\n\t\t\t\t\t\t<br>" +
                        $"\n\t\t\t\t\t\t{ShowBirthdayGivers(true, true)}\n");
                }
            }
            catch
            {
                Configs.AddLogsCollected($"Sending message: FAILURE.");
                /*Console.WriteLine("There are some errors occured, while trying to send message.\nWould you like to reconfigure me? (y/n)");               
                switch (Console.ReadLine().ToLower())
                {
                    case "y":
                    case "yes":
                        CreateConfig();
                        SendMail(args);
                        break;
                    default:
                        break;
                }
                */
            }
        }
        public static void ConfigWriter(string type, string parameter, string value) //TODO: return value.
        {
            string fileType = value.Substring(value.LastIndexOf('.') + 1, value.Length - value.LastIndexOf('.') - 1);
            switch (type)
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
                        value = "";
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
                case "password":
                    break;
                default:
                    break;
            }
            Configs.SetConfigurations(parameter + "=" + value);
            Configs.AddLogsCollected($"Config: " + parameter + "=" + value);
        }

        public static void WriteConfig()
        {            
            Console.Write("Set sender e-mail: ");
            ConfigWriter("text", "senderEmail", Console.ReadLine());

            Console.Write("Set sender password: ");
            ConfigWriter("password", "senderPassword", Console.ReadLine()); //TODO: password enter with mask.

            Console.Write("Set sender displayed name: ");
            ConfigWriter("text", "senderName", Console.ReadLine());            

            Console.Write("Set reciever e-mail: ");
            ConfigWriter("text", "recieverEmail", Console.ReadLine());            

            Console.Write("Set message subject: ");
            ConfigWriter("text", "messageSubject", Console.ReadLine());

            Console.Write($"Set a path to html file: ");
            ConfigWriter("file", "htmlPath", IsFileExist("html")); //Remade IsFileExist

            Configs.SetMessageText(File.ReadAllText(Configs.GetHtmlPath())); //Throw somewhere

            Console.Write($"Set a path to xls file: ");
            ConfigWriter("file", "xlsPath", IsFileExist("xls")); //Remade IsFileExist

            Console.Write("Set a number of column contains birthday dates: ");
            ConfigWriter("digit", "birthdayColumnNumber", IsDigit(false)); //Remade IsDigit            

            Console.Write("Set a number of column contains employees names: ");
            ConfigWriter("digit", "employeeNameColumnNumber", IsDigit(false)); //Remade IsDigit    

            Console.Write("Set server address: ");
            ConfigWriter("text", "serverAddress", Console.ReadLine());

            Console.Write("Set server port (if default - leave empty): ");
            ConfigWriter("digit", "serverPort", IsDigit(true)); //Remade IsDigit     

            Console.WriteLine("Use 5/2 workmode?(yes) \nOtherwise will be user full week mode");
            ConfigWriter("text", "fiveDaysMode", Console.ReadLine()); //TODO: ConfigWriter remade for this. (yes/no)

            Console.Write("Set logs recievers: ");
            ConfigWriter("text", "logRecievers", Console.ReadLine()); //TODO: ConfigWriter remade for this.
            /*
            string recieversString = Console.ReadLine();
            string[] recievers = recieversString.Split(',');
            foreach (var reciever in recievers)
            {
                Configs.SetLogRecievers(reciever.Trim());
            }
            */
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
            ReadXlsFile(Configs.GetXlsPath(), Configs.GetFiveDayMode(), Configs.GetBirthdayColumnNumber(), Configs.GetEmployeeNameColumnNumber());
            ReadHtmlFile(Configs.GetHtmlPath(), Employees.GetCongratulationsString());
            ReadConfig();
            Configs.SetReadConfigSuccess(false);            
        }

        private static string IsFileExist(string fileType)
        {
            string pathToFile;
            while (true)
            {                
                pathToFile = Console.ReadLine();
                if (File.Exists(pathToFile) && pathToFile.EndsWith(fileType))
                {
                    if (File.ReadAllText(pathToFile).Contains("%LIST_OF_EMPLOYEES%") || fileType == "xls")
                    {
                        break;
                    }
                    else
                    {
                        Console.WriteLine("HTML file should have %LIST_OF_EMPLOYEES% string. Current file has no such string.");
                    }
                }
                else
                {
                    Console.WriteLine($"There are no {fileType} file with such name.");
                }
            }
            return pathToFile;
        }

        private static string IsDigit(bool isPort)
        {
            string digit;
            while (true)
            {
                digit = Console.ReadLine();
                if (Int32.TryParse(digit, out _))
                {
                    break;
                }
                else if (isPort && string.IsNullOrEmpty(digit) || string.IsNullOrWhiteSpace(digit))
                {
                    digit = "25";
                    break;
                }
                else
                {
                    Console.WriteLine("Please, use only digits.");
                }
            }
            return digit;
        }

        private static bool IsFileLocked(string filename)
        {
            bool Locked = false;
            if (File.Exists(filename))
            {
                try
                {
                    FileStream fs =
                        File.Open(filename, FileMode.OpenOrCreate,
                        FileAccess.ReadWrite, FileShare.None);
                    fs.Close();
                }
                catch
                {
                    Locked = true;
                }
            }
            return Locked;
        }

        public static void ReadXlsFile(string path, bool fiveDaysMode, string birthdayColumn, string employeeColumn)
        {
            try
            {
                int bdColumn = Convert.ToInt32(birthdayColumn);
                int emColumn = Convert.ToInt32(employeeColumn);
                bdColumn--;
                emColumn--;
                while (IsFileLocked(path))
                {
                    Console.WriteLine("xls file is opened in another application. Please close that app and press any key to continue.");
                    Console.ReadKey();
                }
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
