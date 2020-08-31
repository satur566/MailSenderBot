using System.Collections.Generic;

namespace MailSender
{
    static class Employees
    {
        private static readonly List<string> whosBirthdayIs = new List<string>();
        private static string congratulationsString;

        public static List<string> WhosBirthdayIs
        {
            get
            {
                whosBirthdayIs.Sort();
                return whosBirthdayIs;
            }
            set
            {
                whosBirthdayIs.Clear();
                foreach (string employee in value)
                {
                    whosBirthdayIs.Add(employee);
                }
            }
        }
        public static void AddBirthdaygiver(string entry)
        {
            whosBirthdayIs.Add(entry);
        }
        public static string CongratulationsList
        {
            get
            {
                foreach (var item in whosBirthdayIs)
                {
                    congratulationsString += item.Trim() + "<br>";
                }
                return congratulationsString;
            }
        }
    }
}