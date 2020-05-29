using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace RfLoggerParser
{
    class Program
    {
        static void Main(string[] args)
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            int index = 1;

            string path = args[0];
            List<string> files = new List<string>();
            GetDirectories(path, files);

            foreach (string file in files) {
                string contents = File.ReadAllText(file);
                string imei = Regex.Match(contents, "IMEI :\\s*\\\"(\\d*)").Groups[1].Value;
                string[] tests = Regex.Split(contents, "^\\s*$", RegexOptions.Multiline);

                foreach (string test in tests) {
                    string t = test.Trim();
                    if (Regex.IsMatch(contents, "-* \\w -*")) {
                        string[] info = Regex.Split(t, "-{100}\r?\n");
                        if (info.Length == 4) {
                            string category = Regex.Match(info[0], "-* (.*) -*").Groups[1].Value;
                            string channel = Regex.IsMatch(info[2], "Channel:(\\d*)")? Regex.Match(info[2], "Channel:(\\d*)").Groups[1].Value : "";
                            foreach (string item in Regex.Split(info[3], "\\r?\\n")) {
                                string itemName = item[0..49].Trim();
                                string itemLower = item[50..61].Trim();
                                string itemUpper = item[62..73].Trim();
                                string itemMeasured = item[74..85].Trim();
                                string itemUnit = item[86..91].Trim();
                                string itemStatus = item[92..item.Length].Trim();

                                Console.WriteLine("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}", category, channel, itemName, itemLower, itemUpper, itemMeasured, itemUnit, itemStatus);

                                ws.Cells[index, 1].Value = Path.GetFileName(file);
                                ws.Cells[index, 2].Value = imei;                        
                                ws.Cells[index, 3].Value = category;
                                ws.Cells[index, 4].Value = channel;
                                ws.Cells[index, 5].Value = itemName;
                                ws.Cells[index, 6].Value = itemLower;
                                ws.Cells[index, 7].Value = itemUpper;
                                ws.Cells[index, 8].Value = itemMeasured;
                                ws.Cells[index, 9].Value = itemUnit;
                                ws.Cells[index, 10].Value = itemStatus;
                                index += 1;
                            }                                
                        }
                    }
                }
            }
            p.SaveAs(new FileInfo("./output.xlsx"));
        }

        private static void GetDirectories(string path, List<string> files) {
            foreach (string file in Directory.GetFiles(path)) {
                files.Add(file);
            }

            foreach (string folder in Directory.GetDirectories(path)) {
                GetDirectories(folder, files);
            }
        }
    }
}
