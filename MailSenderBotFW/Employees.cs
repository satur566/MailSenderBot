using System.Collections.Generic;

namespace MailSender
{
    static class Employees
    {
        private static readonly List<string> whosBirthdayIs = new List<string>();
        private static string congratulationsString;

        public static List<string> GetWhosBirthdayIs()
        {
            whosBirthdayIs.Sort();
            return whosBirthdayIs;
        }
        public static void SetWhosBirthdayIs(string entry)
        {
            whosBirthdayIs.Add(entry);
        }

        public static string GetCongratulationsString()
        {
            foreach (var item in whosBirthdayIs)
            {
                congratulationsString += item.Trim() + "<br>";
            }
            return congratulationsString;
        }
    }
}
