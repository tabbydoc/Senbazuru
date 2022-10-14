using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace Senbazuru.HirarchicalExtraction
{
    public class Utils
    {
        public static Dictionary<string, string> NamedParams(string[] args)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            string[] kv = new string[2];

            foreach (string param in args)
            {
                try
                {
                    kv = param.Split('=');
                    result.Add(kv[0], kv[1]);
                }
                catch (Exception e)
                {
                    Debug.WriteLine($"Ошибка: {e.Message}");
                }
            }
            return result;



        }
    }
}
