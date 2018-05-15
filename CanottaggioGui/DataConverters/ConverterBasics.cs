using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CanottaggioGui.DataConverters
{
    public abstract class ConverterBasics
    {
        public Dictionary<string, Tuple<string, string>> Categories { get; set; }
        public Dictionary<string, Tuple<string, string>> Nations { get; set; }
        public Dictionary<string, string> Teams { get; set; }
        protected StringBuilder OutputStream { get; set; }
        public Dictionary<string, string> FileConfig { get; set; }

        public ConverterBasics()
        {
            OutputStream = new StringBuilder();
            LoadBasicFileConfig();
        }

        public abstract bool ConvertInternational(List<Dictionary<string, string>> fields, string title = "");
        public abstract bool ConvertNational(List<Dictionary<string, string>> fields, string title = "");

        public bool Convert(string filepath, bool isNational, string title = "")
        {
            var fields = parseFileExcel(filepath);
            return Convert(fields, isNational, title);
        }
        public bool Convert(List<Dictionary<string, string>> content, bool isNational, string title = "")
        {
            return isNational ? ConvertNational(content, title) : ConvertInternational(content, title);
        }

        public List<Dictionary<string, string>> parseFileExcel(string filepath)
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
                        if (FileConfig.ContainsKey(header_value))
                            header_assoc.Add(columns, header_value);
                        columns++;
                    }
                }
                while (true);
                var listContent = new List<Dictionary<string, string>>();
                for (int row = 2; !string.IsNullOrEmpty(sheet.GetValue(row, 1)?.ToString().Trim()); row++)
                {
                    var rowDictionary = new Dictionary<string, string>();
                    foreach (var column in header_assoc.Keys)
                    {
                        var row_value = sheet.GetValue<string>(row, column);
                        rowDictionary.Add(FileConfig[header_assoc[column]], row_value);
                    }
                    listContent.Add(rowDictionary);
                }
                return listContent.OrderBy(x => Int32.Parse(x["Batteria"])).ThenBy(x => Int32.Parse(x["Acqua"])).ToList();
            }
        }

        protected string GetFlagName(string nation)
        {
            var shortNation = GetShortSurnameInt(nation);
            if (Nations.ContainsKey(nation))
                return Nations[nation].Item1;
            else if (Nations.ContainsKey(shortNation))
                return Nations[shortNation].Item1;
            else
                OutputStream.AppendLine($"Bandiera non trovata: {nation}");
            return nation.Substring(0, 3);
        }
        protected string GetTeamNameInt(string nation)
        {
            var shortNation = GetShortSurnameInt(nation);
            if (Nations.ContainsKey(nation))
                return Nations[nation].Item2.ToUpper();
            else if (Nations.ContainsKey(shortNation))
                return Nations[shortNation].Item2.ToUpper();
            else
                OutputStream.AppendLine($"Team - Nazione ({nation}) non trovata - Short ({nation})");
            return string.Empty;
        }
        protected string GetSurnameInt(string nation, bool isTeam)
        {
            var shortNation = GetShortSurnameInt(nation);
            if (Nations.ContainsKey(nation) && isTeam)
                return Nations[nation].Item2.ToUpper();
            else if (Nations.ContainsKey(shortNation) && isTeam)
                return Nations[shortNation].Item2.ToUpper();
            else if (isTeam)
                OutputStream.AppendLine($"Surname - Nazione ({nation}) non trovata - Short ({shortNation})");
            return shortNation;
        }
        protected string GetShortSurnameInt(string nation)
        {
            var shortNation = nation.ToString();
            if (shortNation.Where(x => x == ' ').Count() >= 2)
            {
                var parts = shortNation.Split(new char[] { ' ' });
                shortNation = $"{parts[0]} {parts[1].Substring(0, 1)}{parts[2]}";
            }
            return shortNation;
        }
        protected string GetCategoryDescription(string catId, string cat2)
        {
            if (Categories.ContainsKey(catId))
                return Categories[catId].Item1;
            else if (Categories.ContainsKey(cat2))
                return Categories[cat2].Item1;
            else
                OutputStream.AppendLine($"Descrizione categoria ({catId}) non trovata. Categoria2 ({cat2})");
            return string.Empty;
        }
        protected string GetCategorySex(string catId, string cat2Id)
        {
            if (Categories.ContainsKey(catId))
                return Categories[catId].Item2;
            else if (Categories.ContainsKey(cat2Id))
                return Categories[cat2Id].Item2;
            else
                OutputStream.AppendLine($"Sex non trovato per la categoria ({catId} - {cat2Id})");
            return string.Empty;
        }
        protected string GetTeamNameNational(string code)
        {
            if (Teams.ContainsKey(code))
                return Teams[code];
            else
                OutputStream.AppendLine($"Squadra non trovata ({code})");
            return string.Empty;
        }
        protected string GetDesktopPath()
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        }
        private void LoadBasicFileConfig()
        {
            FileConfig = new Dictionary<string, string>()
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
        }
        public string GetTextFromStream()
        {
            var text = OutputStream.ToString();
            OutputStream.Clear();
            return text;
        }
    }
}
