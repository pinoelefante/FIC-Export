using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CanottaggioGui.Data
{
    public class AppDataLoader
    {
        private static readonly char[] csvSeparator = new char[] { ';' };
        private static readonly string FILE_TEAMS = @"data\teams.csv";
        private static readonly string FILE_CATEGORIES = @"data\categorie.csv";
        private static readonly string FILE_NATIONS = @"data\nazioni.csv";

        public Dictionary<string, string> LoadFileMappingConfig()
        {
            if (!File.Exists("fileconfig.csv"))
                return null;
            var configLines = File.ReadAllLines("fileconfig.csv");
            Debug.WriteLine($"Ci sono {configLines.Count()} righe nel file di mapping");
            var file_config = new Dictionary<string, string>();
            foreach (var line in configLines)
            {
                var values = line.Split(csvSeparator, StringSplitOptions.None);
                file_config.Add(values[1], values[0]);
            }
            return file_config;
        }
        public Dictionary<string, K> LoadCSV<K>(string file, Func<string[], KeyValuePair<string, K>> func)
        {
            if (!File.Exists(file))
                return new Dictionary<string, K>();
            var lines = File.ReadAllLines(file);
            var content = new Dictionary<string, K>();
            foreach(var line in lines)
            {
                var values = line.Split(csvSeparator, StringSplitOptions.None);
                var x = func.Invoke(values);
                content.Add(x.Key, x.Value);
            }
            return content;
        }
        public Dictionary<string, string> LoadTeamsFile()
        {
            return LoadCSV<string>(FILE_TEAMS, (values) => 
            {
                return new KeyValuePair<string, string>(values[0], values[1]);
            });
        }
        public Dictionary<string, Tuple<string,string>> LoadCategoriesFile()
        {
            return LoadCSV<Tuple<string, string>>(FILE_CATEGORIES, (values) =>
            {
                return new KeyValuePair<string, Tuple<string, string>>(values[0].Trim(), new Tuple<string, string>(values[1].Trim(), values[2].Trim()));
            });
        }
        public Dictionary<string, Tuple<string, string>> LoadNationsFile()
        {
            return LoadCSV(FILE_NATIONS, (values) =>
            {
                return new KeyValuePair<string, Tuple<string, string>>(values[1].Trim(), new Tuple<string, string>(values[0].Trim(), values[2].Trim()));
            });
        }
    }
}
