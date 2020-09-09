using ExcelLibrary.SpreadSheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace MailSender
{
    static class FileWorks
    {
        private static readonly string workingDirectory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
        public static List<string> ReadXlsFile(string path, bool fiveDaysMode, string birthdayColumn, string employeeColumn)
        {
            List<string> employeesList = new List<string>();
            bool isNextMonthLoaded = false;
            bool isTempCopied = false;
            try
            {
                string fileName = path.Substring(path.LastIndexOf('\\'));
                File.Copy(path, TempPath + fileName, true);
                path = TempPath + fileName;
                isTempCopied = true;
            }
            catch
            {
                string exceptionMessage = "Unable to open temporary copy of existing .xls file.";
                Logs.AddLogsCollected(exceptionMessage);
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
                            employeesList.Add(BirthdaySheet.Cells[i, emColumn].ToString());
                        }
                        if (DateTime.Now.DayOfWeek == DayOfWeek.Monday && fiveDaysMode)
                        {
                            try
                            {
                                if (Convert.ToDateTime(BirthdaySheet.Cells[i, bdColumn].ToString()).Date.Equals(DateTime.Now.AddDays(-1).Date))
                                {
                                    employeesList.Add(BirthdaySheet.Cells[i, emColumn].ToString());
                                }
                            }
                            catch { }
                            try
                            {
                                if (Convert.ToDateTime(BirthdaySheet.Cells[i, bdColumn].ToString()).Date.Equals(DateTime.Now.AddDays(-2).Date))
                                {
                                    employeesList.Add(BirthdaySheet.Cells[i, emColumn].ToString());
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
                Logs.AddLogsCollected($"Reading xls: SUCCESS.");
                if (DateTime.Now.Day > 25 && !isNextMonthLoaded)
                {
                    Logs.AddLogsCollected("Xls file does not contain list of employees for a next month.");
                }
                if (isTempCopied)
                {
                    try
                    {
                        File.Delete(path);
                    }
                    catch
                    {
                        string exceptionMessage = "Unable to delete temporary .xls file.";
                        Logs.AddLogsCollected(exceptionMessage);
                        throw new Exception(exceptionMessage);
                    }
                }
            }
            catch
            {
                string exceptionMessage = "Reading xls: FAILURE.";
                Logs.AddLogsCollected(exceptionMessage);
                throw new Exception(exceptionMessage);

            }
            return employeesList;
        }

        public static string ReadHtmlFile(string path, string employees)
        {
            string htmlContent = File.ReadAllText(path);
            if (htmlContent.Contains("%LIST_OF_EMPLOYEES%"))
            {
                htmlContent = htmlContent.Replace("%LIST_OF_EMPLOYEES%", employees);
                Logs.AddLogsCollected($"Reading html: SUCCESS.");                
            }
            else
            {
                Logs.AddLogsCollected($"Reading html: FAILURE.");
                Logs.AddLogsCollected($"Reason: list of employees can't be inserted.");
                htmlContent = null;
            }
            return htmlContent;
        }

        public static List<string> CollectHtmlFiles(string path)
        {
            List<string> filesList = new List<string>();
            foreach (var file in Directory.GetFiles(path, "*.html"))
            {
                string htmlContent = File.ReadAllText(file);
                if (htmlContent.Contains("%LIST_OF_EMPLOYEES%"))
                {
                    filesList.Add(file);
                }
            }
            return filesList;
        }

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
                        string exceptionMessage = "Unable to create logs directory.";
                        Logs.AddLogsCollected(exceptionMessage);
                        throw new Exception(exceptionMessage);
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
                        File.WriteAllLines(configDirectory + "\\config.cfg", Configs.ParametersList);
                    }
                    catch
                    {
                        string exceptionMessage = "Unable to create config directory.";
                        Logs.AddLogsCollected(exceptionMessage);
                        throw new Exception(exceptionMessage);
                    }

                }
                return configDirectory + "\\config.cfg";
            }
        }

        private static string TempPath
        {
            get
            {
                string tempDirectory = workingDirectory + "\\temp";
                if (!Directory.Exists(tempDirectory))
                {
                    try
                    {
                        Directory.CreateDirectory(tempDirectory);
                    }
                    catch
                    {                        
                        string exceptionMessage = "Unable to create temp directory.";
                        Logs.AddLogsCollected(exceptionMessage);
                        throw new Exception(exceptionMessage);
                    }

                }
                return tempDirectory;
            }
        }

        public static void SaveConfig()
        {
            try
            {
                File.WriteAllText(ConfigsPath, string.Empty);
                File.WriteAllLines(ConfigsPath, Configs.ParametersList);
                Logs.AddLogsCollected($"Config save: SUCCESS.");
            }
            catch
            {                
                string exceptionMessage = "Config save: FAILURE.";
                Logs.AddLogsCollected(exceptionMessage);
                throw new Exception(exceptionMessage);
            }
        }

        public static List<string> LoadConfig(string path)
        {
            return new List<string>(File.ReadAllLines(path));            
        }
    }
}
