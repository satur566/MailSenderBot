﻿using System;
using System.IO;

namespace MailSender
{
    static class Logs
    {
        public static void SendLogs()
        {
            foreach (var reciever in Configs.LogsRecievers)
            {
                string ifSuccess = "";
                try
                {
                    Sending.SendMessage(reciever, $"log from {DateTime.Now}", LogsCollected);
                    ifSuccess = "SUCCESS";
                }
                catch
                {
                    ifSuccess = "FAILURE";
                    throw;
                }
                finally
                {
                    AddLogsCollected($"Sending logs to {reciever}: {ifSuccess}.");
                }
            }
        }

        public static string LogsCollected { get; private set; } = "";
        public static string LogConclusionMaker(string reciever)
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

        public static void AddLogsCollected(string log)
        {
            log = $"\n{DateTime.Now} - " + log;
            File.AppendAllText(FileWorks.LogsPath, log);
            LogsCollected = string.Concat(LogsCollected, log.Replace("\t", "&#9;").Replace("\n", "<br>"));
        }
    }
}
