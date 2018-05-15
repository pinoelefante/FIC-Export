using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace CanottaggioConsole
{
    class Program
    {
        private static Dictionary<string, string> file_config = new Dictionary<string, string>()
            {
                { "tbl_pettorale","Pettorale"},
                { "pg_barca&pg_cate&pg_sex", "Categoria2"},
                { "Categoria_int", "Categoria"},
                { "tbl_sex", "Sex"},
                { "tbl_nazione", "Nazione"},
                { "id_squadra", "id_squadra"},
                { "tbl_corsia", "Acqua"},
                { "id_batteria", "Batteria"},
                { "1", "Atleta1"},
                { "2", "Atleta2"},
                { "3", "Atleta3"},
                { "4", "Atleta4"},
                { "5", "Atleta5"},
                { "6", "Atleta6"},
                { "7", "Atleta7"},
                { "8", "Atleta8"},
                { "9", "Atleta9"},
            };
        private static Dictionary<string, Tuple<string, string>> categories = new Dictionary<string, Tuple<string, string>>();
        private static Dictionary<string, Tuple<string, string>> nations = new Dictionary<string, Tuple<string, string>>();
        private static Dictionary<string, string> teams = new Dictionary<string, string>();
        static void Main(string[] args)
        {
            var filename = string.Empty;
            var separator = string.Empty;
            var exportType = string.Empty;
            var national = false;
            var title = string.Empty;
            var base_tvg = string.Empty;
            bool startFromShell = args.Length == 6;
            
            loadFileMapping();
            var contentDictionary = parseFile(filename, separator);
            
        }
        private static void loadFileMapping()
        {
            var configLines = readFile("fileconfig.csv");
            if (configLines == null || configLines.Count() <= 1)
            {
                Console.WriteLine("\nMapping non trovato. Verra' usato il mapping di default");
                return;
            }
            Console.WriteLine($"Ci sono {configLines.Count()} righe nel file di mapping");
            file_config.Clear();
            var splitVal = new char[] { ';' };
            foreach(var line in configLines)
            {
                var values = line.Split(splitVal, StringSplitOptions.None);
                file_config.Add(values[1], values[0]);
            }
        }
        private static string getStringFromCommandLine(string text, bool lower=true, bool acceptAll = false, params string[] validValues)
        {
            var line = string.Empty;
            do
            {
                Console.Write(text);
                line = Console.ReadLine().Trim();
                if (lower)
                    line = line.ToLower();
            } while (!isValidValue(line, validValues) && !acceptAll);
            return line;
        }
        private static string getStringAsPathFromCommandLine(string text, bool file = true)
        {
            var line = string.Empty;
            do
            {
                Console.Write(text);
                line = Console.ReadLine().Trim();
                if (file && File.Exists(line))
                    break;
                if (!file && Directory.Exists(line))
                    break;
            } while (true);
            return line;
        }
        private static bool getStringAsBoolFromCommandLine(string text, bool lower, string[] positiveValues, string[] negativeValues)
        {
            var line = string.Empty;
            do
            {
                Console.Write(text);
                line = Console.ReadLine().Trim();
                if (lower)
                    line = line.ToLower();
                if (isValidValue(line, positiveValues))
                    return true;
                if (isValidValue(line, negativeValues))
                    return false;
            } while (true);
        }
        private static bool isValidValue(string input, string[] validValues)
        {
            foreach (var vValue in validValues)
                if (vValue.CompareTo(input) == 0)
                    return true;
            return false;
        }
        private static void loadCategories()
        {
            categories.Clear();
            var lines = File.ReadAllLines(@"data\categorie.csv");
            var splitVal = new char[] { ';' };
            foreach (var line in lines)
            {
                var values = line.Split(splitVal, StringSplitOptions.None);
                categories.Add(values[0].Trim(), new Tuple<string, string>(values[1].Trim(), values[2].Trim()));
            }
        }
        private static void loadNations()
        {
            nations.Clear();
            var lines = File.ReadAllLines(@"data\nazioni.csv");
            var splitVal = new char[] { ';' };
            foreach(var line in lines)
            {
                var values = line.Split(splitVal, StringSplitOptions.None);
                nations.Add(values[1].Trim(), new Tuple<string, string>(values[0].Trim(), values[2].Trim()));
            }
        }
        private static void loadTeams()
        {
            teams.Clear();
            var lines = File.ReadAllLines(@"data\teams.csv");
            var splitVal = new char[] { ';' };
            foreach(var line in lines)
            {
                var values = line.Split(splitVal, StringSplitOptions.None);
                teams.Add(values[0], values[1]);
            }
        }
        private static string[] readFile(string path)
        {
            if(File.Exists(path))
                return File.ReadAllLines(path);
            return null;
        }
        private static List<Dictionary<string,string>> parseFile(string filepath, string separator=";")
        {
            if (!File.Exists(filepath))
                throw new FileNotFoundException($"Il file {filepath} non esiste");
            if (!filepath.Contains('.'))
                throw new Exception("Il file non ha un'estensione");

            var fileExt = filepath.Substring(filepath.LastIndexOf('.') + 1).ToLower();
            switch(fileExt)
            {
                case "csv":
                    var fileContent = readFile(filepath);
                    return parseFileCSV(fileContent, separator);
                case "xlsx":
                    return parseFileExcel(filepath);
                default:
                    throw new Exception("Tipo di file non supportato");
            }
        }
        private static List<Dictionary<string, string>> parseFileExcel(string filepath)
        {
            using (var excelFile = new ExcelPackage(new FileInfo(filepath)))
            {
                var sheet = excelFile.Workbook.Worksheets.First();
                Dictionary<int, string> header_assoc = new Dictionary<int, string>();
                int columns = 1;
                do
                {
                    var header_value = sheet.GetValue<string>(1, columns);
                    if (string.IsNullOrEmpty(header_value))
                        break;
                    else
                    {
                        if (file_config.ContainsKey(header_value))
                            header_assoc.Add(columns, header_value);
                        columns++;
                    }
                }
                while (true);
                var listContent = new List<Dictionary<string, string>>();
                for(int row = 2; !string.IsNullOrEmpty(sheet.GetValue(row, 1)?.ToString().Trim()); row++)
                {
                    var rowDictionary = new Dictionary<string, string>();
                    foreach (var column in header_assoc.Keys)
                    {
                        var row_value = sheet.GetValue<string>(row, column);
                        rowDictionary.Add(file_config[header_assoc[column]], row_value);
                    }   
                    listContent.Add(rowDictionary);
                }
                return listContent;
            }
        }
        private static List<Dictionary<string,string>> parseFileCSV(string[] content, string separator)
        {
            var separatorPar = separator.ToCharArray();
            var headers = content[0].Split(separatorPar, StringSplitOptions.None);
            Dictionary<int, string> header_assoc = new Dictionary<int, string>();
            for (int i = 0;i<headers.Length;i++)
            {
                var h_value = headers[i];
                if(file_config.ContainsKey(h_value))
                    header_assoc.Add(i, h_value);
            }
            List<Dictionary<string, string>> listContent = new List<Dictionary<string, string>>(content.Length);
            for(int i=1;i<content.Length;i++)
            {
                var rowContent = content[i].Split(separatorPar, StringSplitOptions.None);
                var rowDictionary = new Dictionary<string, string>();
                foreach(var index in header_assoc.Keys)
                    rowDictionary.Add(file_config[header_assoc[index]], rowContent[index]);
                listContent.Add(rowDictionary);
            }
            return listContent.OrderBy(x => Int32.Parse(x["Batteria"])).ThenBy(x=> Int32.Parse(x["Acqua"])).ToList();
        }
    }
    
}
