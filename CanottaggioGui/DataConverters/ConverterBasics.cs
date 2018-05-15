using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CanottaggioConsole.DataConverters
{
    public abstract class ConverterBasics
    {
        public Dictionary<string, Tuple<string, string>> Categories { get; set; }
        public Dictionary<string, Tuple<string, string>> Nations { get; set; }
        public Dictionary<string, string> Teams { get; set; }
        public StreamWriter OutputStream { get; set; }

        public ConverterBasics()
        {
            OutputStream = new StreamWriter(new MemoryStream());
        }

        public bool Convert(List<Dictionary<string, string>> fields, bool isNational, string title = "")
        {
            return isNational ? ConvertNational(fields) : ConvertInternational(fields);
        }
        public abstract bool ConvertInternational(List<Dictionary<string, string>> fields, string title="");
        public abstract bool ConvertNational(List<Dictionary<string, string>> fields, string title = "");

        protected string GetFlagName(string nation)
        {
            var shortNation = GetShortSurnameInt(nation);
            if (Nations.ContainsKey(nation))
                return Nations[nation].Item1;
            else if (Nations.ContainsKey(shortNation))
                return Nations[shortNation].Item1;
            else
                OutputStream.WriteLine($"Bandiera non trovata: {nation}");
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
                OutputStream.WriteLine($"Team - Nazione ({nation}) non trovata - Short ({nation})");
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
                OutputStream.WriteLine($"Surname - Nazione ({nation}) non trovata - Short ({shortNation})");
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
                OutputStream.WriteLine($"Descrizione categoria ({catId}) non trovata. Categoria2 ({cat2})");
            return string.Empty;
        }
        protected string GetCategorySex(string catId, string cat2Id)
        {
            if (Categories.ContainsKey(catId))
                return Categories[catId].Item2;
            else if (Categories.ContainsKey(cat2Id))
                return Categories[cat2Id].Item2;
            else
                OutputStream.WriteLine($"Sex non trovato per la categoria ({catId} - {cat2Id})");
            return string.Empty;
        }
        protected string GetTeamNameNational(string code)
        {
            if (Teams.ContainsKey(code))
                return Teams[code];
            else
                OutputStream.WriteLine($"Squadra non trovata ({code})");
            return string.Empty;
        }
        protected string GetDesktopPath()
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        }
    }
}
