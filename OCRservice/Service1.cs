using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace OCRservice
{
    public partial class Service1 : ServiceBase
    {
        FileSystemWatcher watcher;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteLog("Service is started at " + DateTime.Now);
            CreateFileWatcher(@"C:\Dev");
            Console.ReadLine();
        }

        protected override void OnStop()
        {
            watcher.EnableRaisingEvents = true;
            WriteLog("Service Stopped at " + DateTime.Now);
        }

        public static void WriteLog(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }

        static void ParseFile()
        {
            var lines = File.ReadAllLines(@"C:\Dev\OCRlog.log");
            foreach (var line in lines)
            {
                CreateHtml(line);
            }
        }

        static void CreateHtml(string content)
        {
            var stringArray = content.Split(';');

            var first_name = stringArray[2];
            var last_name = stringArray[3];
            var ic_number = stringArray[4];
            var birth_date = stringArray[6];
            var sex = stringArray[7];
            var scaned_at = stringArray[0];
            var valid_to = stringArray[5];
            var ic_type = stringArray[12];
            var country_code = stringArray[9];
            string htmlContent = $@"<b><h1>Hotel Antik</h1></b>Dlouhá 22, Praha 1, 110 00, Czech Republic
                <table border=1 cellpadding=20 cellspacing=0 width=100% rules=groups>
                <colgroup span=4 align=righ2>
                <col align=left width=6%>
                <col align=center width=25%>
                <col align=right width=5%>
                <col align=right2 width=25%>
                <td>Family Name: <br>Příjmení:<td><b><h1>
                {last_name}
                </b></h1></td><td>First Name: <br>Jméno:<td><b><h1>
                {first_name}
                </b></h1><tr><td>Date of Birth: <br>Datum narození:<td><b><h2>
                {birth_date}
                </b></h1></td><td>Citizenship: <br> Státní občanství:<td><b><h2>
                {country_code} 
                </b></h2></td>
                <tr><td>ID No.: <br>Číslo dokladu:<td><b><h2>
                {ic_number}
                </b></h2></td><td>ID Type: <br>Typ dokladu:<td><b><h2>
                {ic_type}
                </b></h2></td>
                <tr><td>Valid No.: <br>Platnost dokladu:<td><b><h2>
                {valid_to}
                </b></h2></td><td>Scaned at: <br>Načteno v:<td><h4>
                {scaned_at}
                </h4></td>
                </table>
                Souhlasím / I agree <br> Se zpracováním osobních údajů pro marketingové účely / 
                With the processing of personal data for marketing purposes. <br>Veškeré údaje jsou vyžadovány dle zák.č. 565/1990 a 290/2004 Sb. a je s nimi nakládáno v souladu se zák.č. 101/2000 Sb. / 
                All information is necessary for our law and all information will be used in compliance with our law.
                <br>Ubytovaný host prohlašuje, že byl seznámen s ubytovacím řádem v dostatečném předstihu před poskytnutím služby. Orgánem vykonávajícím dohled nad ochranou spotřebitele a subjektem mimosoudního řešení spotřebitelských sporů je
                <br> Česká obchodní inspekce, Štěpánská 15, Praha 2, 12000, email: adr@coi.cz, www.coi.cz <br><b>Děkujeme / Thank you</b>";

            if (ic_type != "Credit_Card")
            {
                File.WriteAllText(@"C:\Dev\OCR.html", htmlContent);
            }

            File.WriteAllText(@"C:\Dev\OCRlog.log", "");
        }

        public static void CreateFileWatcher(string path)
        {
            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = path;
            watcher.NotifyFilter = NotifyFilters.LastWrite;
            watcher.Filter = "*.log";

            watcher.Changed += new FileSystemEventHandler(OnChanged);
            watcher.Created += new FileSystemEventHandler(OnChanged);

            watcher.EnableRaisingEvents = true;
            WriteLog("FileWatcher registered at " + DateTime.Now);
        }

        private static void OnChanged(object source, FileSystemEventArgs e)
        {
            WriteLog("File: " + e.FullPath + " " + e.ChangeType + DateTime.Now);
            ParseFile();
        }
    }
}
