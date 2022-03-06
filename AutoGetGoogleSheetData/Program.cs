using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Net;

namespace AutoGetGoogleSheetData
{
    class Program
    {
        private const string KEY = "AIzaSyAkpLxtgUPyHYDRLZlTc1veBhcW1WDhqnY";
        private const string GET_DATA_URL = "http://172-105-200-91.ip.linodeusercontent.com:8080/api/v1/sheetData/?id={0}&name={1}&pretty=1";


        private class GetDataSetting
        {
            public string sheetName;
            public string sheetID;
        }

        [STAThread]
        static void Main(string[] args)
        {
            string _path = Directory.GetCurrentDirectory();
            string _settingPath = _path + "/Setting.txt";
            bool _exist = File.Exists(_settingPath);

            Console.WriteLine("EXE path=" + _path);
            Console.WriteLine("Setting exits=" + _exist);

            if (!_exist)
            {
                Console.WriteLine("----------------------------");
                Console.WriteLine("Can't find Setting file, please check.");
                Console.WriteLine("----------------------------");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("----------------------------");

            string _rawSettingData = File.ReadAllText(_settingPath);
            string[] _rawSettingDataParts = _rawSettingData.Split('\n');

            List<GetDataSetting> _getDataSettings = new List<GetDataSetting>();

            for (int i = 0; i < _rawSettingDataParts.Length; i++)
            {
                if (string.IsNullOrEmpty(_rawSettingDataParts[i]))
                {
                    continue;
                }

                string[] _settingParts = _rawSettingDataParts[i].Split('|');

                GetDataSetting _newSettingData = new GetDataSetting
                {
                    sheetName = _settingParts[0].Trim(),
                    sheetID = _settingParts[1].Trim()
                };

                _getDataSettings.Add(_newSettingData);
                Console.WriteLine("Set GetDataSetting: sheetName=" + _newSettingData.sheetName + ", sheetID=" + _newSettingData.sheetID);
            }

            Console.WriteLine("----------------------------");

            string _outputPath = _path + "/Output";

            if (!Directory.Exists(_outputPath))
            {
                Console.WriteLine("output folder didn't exist, creating...");
                Directory.CreateDirectory(_outputPath);
            }

            for (int i = 0; i < _getDataSettings.Count; i++)
            {
                string _url = string.Format(GET_DATA_URL, _getDataSettings[i].sheetID, _getDataSettings[i].sheetName, KEY);
                string result = "";

                Console.WriteLine(_url);

                HttpWebRequest request = HttpWebRequest.Create(_url) as HttpWebRequest;
                request.Method = "GET";
                request.ContentType = "application/x-www-form-urlencoded";
                request.Timeout = 30000;

                using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
                {
                    using (StreamReader sr = new StreamReader(response.GetResponseStream()))
                    {
                        result = sr.ReadToEnd();
                    }
                }

                string _outputFilePath = _outputPath + "/" + _getDataSettings[i].sheetName + ".txt";
                if (File.Exists(_outputFilePath))
                {
                    File.WriteAllText(_outputFilePath, result);
                }
                else
                {
                    File.Create(_outputFilePath).Dispose();
                    File.WriteAllText(_outputFilePath, result);
                }

                Console.WriteLine("Text file is set");
            }

            Console.WriteLine("---End---");

            Console.WriteLine("Please Press Any Key");

            Console.ReadKey();
        }

        private static void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser webBrowser = (WebBrowser)sender;
            Console.WriteLine(webBrowser.Document.GetElementsByTagName("span")[0].InnerHtml);
        }

        private static void WaitTillLoad(WebBrowser webBrowser)
        {
            WebBrowserReadyState loadStatus = default(WebBrowserReadyState);
            int waittime = 100000;
            int counter = 0;
            while (true)
            {
                loadStatus = webBrowser.ReadyState;
                Application.DoEvents();

                if ((counter > waittime) || (loadStatus == WebBrowserReadyState.Uninitialized) || (loadStatus == WebBrowserReadyState.Loading) || (loadStatus == WebBrowserReadyState.Interactive))
                {
                    break; // TODO: might not be correct. Was : Exit While
                }
                counter += 1;
            }
            counter = 0;
            while (true)
            {
                loadStatus = webBrowser.ReadyState;
                Application.DoEvents();

                if (loadStatus == WebBrowserReadyState.Complete)
                {
                    break; // TODO: might not be correct. Was : Exit While
                }

                counter += 1;
            }
        }
    }
}
