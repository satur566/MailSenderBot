using System;
using System.IO;

namespace MailSender
{
    static class Logs
    {
        public static void SendLogs()
        {
            foreach (var reciever in Configs.LogsRecievers)
            {
                Sending.SendMessage(reciever, $"log from {DateTime.Now}", LogsCollected);
                AddLogsCollected($"Sending logs to {reciever}: SUCCESS.");
            }
        }

        public static string LogsCollected { get; private set; } = "";
        public static string LogConclusionMaker(string reciever)
        {
            string employees = "";
            foreach (var item in Employees.WhosBirthdayIs)
            {
                employees = string.Concat(employees, "\t" + item.Trim() + "\n");

            }
            string result = $"\nConclusion:" +
                $"\nSender mail: {Configs.SenderEmail}" +
                $"\nSender name: {Configs.SenderName}" +
                $"\nReciever e-mail:{reciever}" +
                $"\nBirthday givers:\n{employees}";
            return result;
        }

        public static void AddLogsCollected(string log)
        {
            log = $"\n{DateTime.Now} - " + log;
            LogsCollected = string.Concat(LogsCollected, log.Replace("\t", "&#9;").Replace("\n", "<br>"));
            try
            {
                File.AppendAllText(FileWorks.LogsPath, log);
            }
            catch
            {
                string message = $"\nUnable to write logs in {FileWorks.LogsPath}";
                LogsCollected = string.Concat(LogsCollected, message.Replace("\t", "&#9;").Replace("\n", "<br>"));
            }            
        }
    }
}
