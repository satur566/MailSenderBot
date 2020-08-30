using ExcelLibrary.SpreadSheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailSender
{
    static class FileReader
    {
        public static List<string> ReadXlsFile(string path, bool fiveDaysMode, string birthdayColumn, string employeeColumn)
        {
            List<string> employeesList = new List<string>();
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
                Logs.AddLogsCollected("Unable to open temporary copy of existing .xls file.");
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
                    try
                    {

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
                        Logs.AddLogsCollected("Unable to delete temporary .xls file.");
                    }
                }
            }
            catch
            {
                Logs.AddLogsCollected($"Reading xls: FAILURE.");
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
            }
            return htmlContent;
        }
    }
}
