using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExportXLS
{


    static class Logger
    {
        static string fName = Path.Combine(Directory.GetCurrentDirectory(), "Export.log");

        public static void Log(string message)
        {
            StreamWriter log = new StreamWriter(fName, true);
            log.WriteLine(DateTime.Now.ToString() + ": " + message);
            log.Close();
        }

    }
}
