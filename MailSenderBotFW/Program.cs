using System;
using System.Net.Mail;
using System.Net;
using System.IO;
using ExcelLibrary.SpreadSheet;
using System.Collections.Generic;
using System.Linq;

namespace MailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            Configs.AddLogsCollected($"\n\n\nCurrent user: {Environment.UserName}");
            try
            {
                Console.ReadKey();
                ReadConfig();
                Configs.SetReadConfigSuccess(true);
                Configs.AddLogsCollected($"Loading config: SUCCESS");

            }
            catch
            {
                Console.WriteLine("Cannot properly read config file. Please set configuration again.");
                Configs.AddLogsCollected($"Loading config: FAILURE");
                CreateConfig();
                Configs.SetReadConfigSuccess(false);
            }
            try
            {
                if (DateTime.Now.Day == 2 && DateTime.Now.Month == 8)
                {
                    SendMessage("a.maksimov@sever.ttk.ru", "Happy Birthday!", "Happy birthday, daddy! Wish you a good incoming year!", args, false);
                    SendMessage("satur566@gmail.com", "Happy Birthday!", "Happy birthday, daddy! Wish you a good incoming year!", args, false);
                }
            }
            catch { }
            PresendCheck(args);
        }

        private static void PresendCheck(string[] args)
        {
            if (String.IsNullOrEmpty(ShowBirthdayGivers(false, false)) || Configs.GetFiveDayMode() && (DateTime.Now.DayOfWeek == DayOfWeek.Sunday || DateTime.Now.DayOfWeek == DayOfWeek.Saturday))
            {
                Console.WriteLine("Today is day off");
                if (String.IsNullOrEmpty(ShowBirthdayGivers(false, false)))
                {
                    Configs.AddLogsCollected($"Sending message: CANCELLED.");
                    Configs.AddLogsCollected($"Reason: employees don't have a birthday today.");
                    SendLogs(args);
                }
                else
                {
                    Configs.AddLogsCollected($"Sending message: CANCELLED.");
                    Configs.AddLogsCollected($"Reason: today is a day off.");
                    SendLogs(args);
                }
            }
            else
            {
                Console.Write($"Conclusion:" +
                    $"\nSender mail: {Configs.GetSenderEmail()}" +
                    $"\nSender name: {Configs.GetSenderName()}" +
                    $"\nReciever e-mail: {Configs.GetRecieverEmail()}" +
                    $"\n" +
                    $"\n{ShowBirthdayGivers(false, false)}\n");

                if (args.Contains<string>("-silent") && Configs.GetReadConfigSuccess())
                {
                    SendMessage(Configs.GetRecieverEmail(), Configs.GetMessageSubject(), Configs.GetMessageText(), args, true);
                    Configs.AddLogsCollected($"Sending message mode: silent");
                    SendLogs(args);
                }
                else
                {
                    Console.WriteLine("\nIs everything fine? (Y/N)");
                    switch (Console.ReadLine().ToLower())
                    {
                        case "y":
                            SendMessage(Configs.GetRecieverEmail(), Configs.GetMessageSubject(), Configs.GetMessageText(), args, true);
                            SendLogs(args);
                            break;
                        default:
                            Configs.AddLogsCollected($"Sending message: CANCELLED.");
                            Configs.AddLogsCollected($"Reason: user cancel.");
                            SendLogs(args);
                            break;
                    }
                }
            }
        }

        private static void SendLogs(string[] args)
        {
            foreach (var reciever in Configs.GetLogRecievers())
            {
                try
                {
                    SendMessage(reciever, $"log from {DateTime.Now}", Configs.GetLogsCollected(), args, false);
                }
                catch
                {

                }
            }
        }

        private static string ShowBirthdayGivers(bool inLogs, bool onEmail)
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

        private static void ReadConfig()
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
            ReadFile(Configs.GetXlsPath(), Configs.GetFiveDayMode(), Configs.GetBirthdayColumnNumber(), Configs.GetEmployeeNameColumnNumber());
            CheckHtml();
        }

        private static void CheckHtml()
        {
            if (File.ReadAllText(Configs.GetHtmlPath()).Contains("%LIST_OF_EMPLOYEES%"))
            {
                Configs.SetMessageText(File.ReadAllText(Configs.GetHtmlPath()).Replace("%LIST_OF_EMPLOYEES%", Employees.GetCongratulationsString()));
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
                    Configs.AddLogsCollected($"Sending message: SUCCESS.");                    
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
                Console.WriteLine("There are some errors occured, while trying to send message.\nWould you like to reconfigure me? (y/n)");
                Configs.AddLogsCollected($"Sending message: FAILURE.");
                switch (Console.ReadLine().ToLower())
                {
                    case "y":
                    case "yes":
                        CreateConfig();
                        PresendCheck(args);
                        break;
                    default:
                        break;
                }

            }
        }

        private static void CreateConfig()
        {
            Console.Write("Set sender e-mail: ");
            Configs.SetSenderEmail(Console.ReadLine());
            Configs.SetConfigurations("senderEmail=" + Configs.GetSenderEmail());
            Configs.AddLogsCollected($"Config: senderEmail={Configs.GetSenderEmail()}");

            Console.Write("Set sender password: ");
            Configs.SetSenderPassword(Console.ReadLine()); //TODO: password enter with mask.
            Configs.SetConfigurations("senderPassword=" + Configs.GetSenderPassword());
            Configs.AddLogsCollected($"Config: senderPassword={Configs.GetSenderPassword()}");

            Console.Write("Set sender displayed name: ");
            Configs.SetSenderName(Console.ReadLine());
            Configs.SetSenderUsername(Configs.GetSenderEmail());
            Configs.SetConfigurations("senderName=" + Configs.GetSenderName());
            Configs.AddLogsCollected($"Config: senderName={Configs.GetSenderName()}");
            Configs.SetConfigurations("senderUsername=" + Configs.GetSenderUsername());
            Configs.AddLogsCollected($"Config: senderUsername={Configs.GetSenderUsername()}");

            Console.Write("Set reciever e-mail: ");
            Configs.SetRecieverEmail(Console.ReadLine());
            Configs.SetConfigurations("recieverEmail=" + Configs.GetRecieverEmail());
            Configs.AddLogsCollected($"Config: recieverEmail={Configs.GetRecieverEmail()}");

            Console.Write("Set message subject: ");
            Configs.SetMessageSubject(Console.ReadLine());
            Configs.SetConfigurations("messageSubject=" + Configs.GetMessageSubject());
            Configs.AddLogsCollected($"Config: messageSubject={Configs.GetMessageSubject()}");

            Configs.SetHtmlPath(IsFileExist("html"));
            Configs.SetConfigurations("htmlPath=" + Configs.GetHtmlPath());
            Configs.AddLogsCollected($"Config: htmlPath={Configs.GetHtmlPath()}");

            Configs.SetMessageText(File.ReadAllText(Configs.GetHtmlPath()));

            Configs.SetXlsPath(IsFileExist("xls"));
            Configs.SetConfigurations("xlsPath=" + Configs.GetXlsPath());
            Configs.AddLogsCollected($"Config: xlsPath={Configs.GetXlsPath()}");

            Console.Write("Set a number of column contains birthday dates: ");
            Configs.SetBirthdayColumnNumber(IsDigit(false));
            Configs.SetConfigurations("birthdayColumnNumber=" + Configs.GetBirthdayColumnNumber());
            Configs.AddLogsCollected($"Config: birthdayColumnNumber={Configs.GetBirthdayColumnNumber()}");

            Console.Write("Set a number of column contains employees names: ");
            Configs.SetEmployeeNameColumnNumber(IsDigit(false));
            Configs.SetConfigurations("employeeNameColumnNumber=" + Configs.GetEmployeeNameColumnNumber());
            Configs.AddLogsCollected($"Config: employeeNameColumnNumber={Configs.GetEmployeeNameColumnNumber()}");

            Console.Write("Set server address: ");
            Configs.SetServerAddress(Console.ReadLine());
            Configs.SetConfigurations("serverAddress=" + Configs.GetServerAddress());
            Configs.AddLogsCollected($"Config: serverAddress={Configs.GetServerAddress()}");

            Console.Write("Set server port (if default - leave empty): ");
            Configs.SetServerPort(IsDigit(true));
            Configs.SetConfigurations("serverPort=" + Configs.GetServerPort());
            Configs.AddLogsCollected($"Config: serverPort={Configs.GetServerPort()}");

            Console.WriteLine("Use 5/2 workmode?(yes) \nOtherwise will be user full week mode");
            switch (Console.ReadLine().ToLower())
            {
                case "y":
                case "yes":
                    Configs.SetFiveDayMode(true);
                    break;
                default:
                    Configs.SetFiveDayMode(false);
                    break;
            }
            Configs.SetConfigurations("fiveDaysMode=" + Configs.GetFiveDayMode());
            Configs.AddLogsCollected($"Config: fiveDaysMode={Configs.GetFiveDayMode()}");

            Console.Write("Set logs recievers: ");
            string recieversString = Console.ReadLine();
            string[] recievers = recieversString.Split(',');
            foreach (var reciever in recievers)
            {
                Configs.SetLogRecievers(reciever.Trim());
            }
            Configs.SetConfigurations("logRecievers=" + recieversString);
            Configs.AddLogsCollected("logRecievers=" + recieversString);

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
            try
            {
                ReadFile(Configs.GetXlsPath(), Configs.GetFiveDayMode(), Configs.GetBirthdayColumnNumber(), Configs.GetEmployeeNameColumnNumber());
                Configs.AddLogsCollected($"Reading xls: SUCCESS.");
            }
            catch
            {
                Configs.AddLogsCollected($"Reading xls: FAILURE.");
            }

            CheckHtml();
        }

        private static string IsFileExist(string fileType)
        {
            string pathToFile;
            while (true)
            {
                Console.Write($"Set a path to {fileType} file: ");
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

        static bool IsFileLocked(string filename)
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

        static void ReadFile(string path, bool fiveDaysMode, string birthdayColumn, string employeeColumn)
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
        }
    }
}
