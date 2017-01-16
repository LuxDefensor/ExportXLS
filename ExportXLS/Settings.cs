using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExportXLS
{
    class Settings
    {
        private string server;
        private string database;
        private string userName;
        private string password;
        private string sMTPServer;
        private string sMTPUserName;
        private string sMTPPassword;
        private int sMTPPort;
        private string addressFrom;
        private int cleanUp;
        private int useSSL;

        #region Properties

        public string Server
        {
            get
            {
                return server;
            }
        }

        public string Database
        {
            get
            {
                return database;
            }
        }

        public string UserName
        {
            get
            {
                return userName;
            }
        }

        public string Password
        {
            get
            {
                return password;
            }
        }

        public string SMTPServer
        {
            get
            {
                return sMTPServer;
            }
        }

        public string SMTPUserName
        {
            get
            {
                return sMTPUserName;
            }
        }

        public string SMTPPassword
        {
            get
            {
                return sMTPPassword;
            }
        }

        public int SMTPPort
        {
            get
            {
                return sMTPPort;
            }
        }

        public string AddressFrom
        {
            get
            {
                return addressFrom;
            }
        }

        public int CleanUp
        {
            get
            {
                return cleanUp;
            }
        }

        public int UseSSL
        {
            get
            {
                return useSSL;
            }
        }
        #endregion

        public Settings()
        {
            string[] lines = File.ReadAllLines("Settings.ini");
            Dictionary<string, string> s =
                lines.ToDictionary<string, string, string>((string inp) => inp.Split(new char[] { '=', ' ' }, StringSplitOptions.RemoveEmptyEntries)[0].Trim(),
                                                           (string el) => el.Split(new char[] { '=', ' ' }, StringSplitOptions.RemoveEmptyEntries)[1].Trim());
            try
            {
                server = s["Server"];
                database = s["Database"];
                userName = s["UserName"];
                password = s["Password"];
                sMTPServer = s["SMTPServer"];
                sMTPPort = int.Parse(s["SMTPPort"]);
                sMTPUserName = s["SMTPUserName"];
                sMTPPassword = s["SMTPPassword"];
                addressFrom = s["AddressFrom"];
                cleanUp =int.Parse(s["CleanUp"]);
                useSSL = int.Parse(s["UseSSL"]);
            }
            catch (Exception ex)
            {
                throw new Exception("Не могу прочитать настройки.", ex);
            }
        }

    }
}