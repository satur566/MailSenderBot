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
                Methods.ReadConfig();                                
            }
            catch
            {
                Console.WriteLine("Cannot properly read config file. Please set configuration again.");                
                Methods.WriteConfig();                
            }
            if (Configs.GetReadConfigSuccess())
            {
                Configs.AddLogsCollected($"Loading config: SUCCESS");
            } 
            else
            {
                Configs.AddLogsCollected($"Loading config: FAILURE");
            }
            try
            {
                /*if (DateTime.Now.Day == 2 && DateTime.Now.Month == 8)
                {
                    Methods.SendMessage("a.maksimov@sever.ttk.ru", "Happy Birthday!", "Happy birthday, daddy! Wish you a good incoming year!", args, false);
                    Methods.SendMessage("satur566@gmail.com", "Happy Birthday!", "Happy birthday, daddy! Wish you a good incoming year!", args, false);
                }*/
            }
            catch { }
            Methods.SendMail(args);
        }       
    }
}
