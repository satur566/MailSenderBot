using System;
using System.IO;
using System.Linq;

namespace MailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            Configs.AddLogsCollected($"\n\n\nCurrent user: {Environment.UserName}"); //TODO: editConfig with args -edit parameter1=value1 parameter2=value2   //TODO: -nopassword parameter //-silent parameter
            if (args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i].StartsWith("-"))
                    {
                        switch (args[i].ToLower()) {
                            case "-help":
                                //TODO: describe every parameter.
                                break;
                            case "-showconfig":
                                foreach (var line in File.ReadAllLines(Configs.GetConfigPath())) 
                                {
                                    Console.WriteLine(line);
                                }
                                break;
                            case "-editconfig":
                                if (i + 1 <= args.Length) { //TODO: trouble with LoadConfig(). Multiple usage cause problems. Needs to be solved. Troubles: multiplicates mail and logs recievers, employees list.
                                    try
                                    {
                                        Methods.LoadConfig();
                                    }
                                    catch
                                    {
                                        Console.WriteLine("Unable to find previous configuration.");
                                    }
                                    for (int j = i + 1; j < args.Length && !args[j].StartsWith("-"); j++)
                                    {
                                        try
                                        {
                                            string parameter = args[j].Substring(0, args[j].IndexOf('='));
                                            string  value = args[j].Substring(args[j].IndexOf('=') + 1, args[j].Length - args[j].IndexOf('=') - 1);                                               
                                            Methods.EditConfig(parameter, value);
                                        }
                                        catch
                                        {
                                            Console.WriteLine("Unable to edit configuration. Invalid parameter.");
                                        }
                                        i++;
                                    }
                                }
                                Methods.SaveConfig();
                                break;
                            default:
                                Console.WriteLine("Unknown parameter.");
                                break;
                        }
                    }
                }
                if(args.Contains("-silent"))
                {
                    try
                    {
                        Configs.AddLogsCollected("Working mode: silent");
                        Methods.LoadConfig();
                        Methods.SendMail();
                    }
                    catch
                    {
                        Console.WriteLine("Unable to send message. Check configuration.");
                    }
                }
            } else
            {
                //TODO: launch WPF app.
            }                       
        }
    }
}
