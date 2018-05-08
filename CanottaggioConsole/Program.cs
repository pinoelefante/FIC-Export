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
        static void Main(string[] args)
        {
            var filename = string.Empty;
            var separator = string.Empty;
            var exportType = string.Empty;
            var national = false;

            credits();

            if (args.Length != 4)
            {
                do
                {
                    Console.Write("Percorso file: ");
                    filename = Console.ReadLine().Trim();
                    if (File.Exists(filename))
                        break;
                    else
                        Console.WriteLine("File inesistente");
                } while (true);
                
                Console.Write("Separatore CSV [default = ;]: ");
                separator = Console.ReadLine().Trim();
                if(string.IsNullOrEmpty(separator))
                {
                    separator = ";";
                    Console.WriteLine("Verra' utilizzato il carattere ; come separatore CSV");
                }

                do
                {
                    Console.Write("Export [mispeaker/tvg]: ");
                    exportType = Console.ReadLine().Trim().ToLower();
                    if (exportType.CompareTo("mispeaker") == 0 || exportType.CompareTo("tvg") == 0)
                        break;
                    else
                        Console.WriteLine("Valore non valido");
                } while (true);

                do
                {
                    Console.Write("Nazionale [true/false]: ");
                    var consoleValue = Console.ReadLine().Trim().ToLower();
                    switch(consoleValue)
                    {
                        case "t":
                        case "true":
                        case "v":
                        case "vero":
                        case "s":
                        case "si":
                        case "nazionale":
                            consoleValue = "true";
                            break;
                        case "internazionale":
                        case "n":
                        case "no":
                        case "f":
                        case "false":
                        case "falso":
                            consoleValue = "false";
                            break;
                        default:
                            continue;
                    }
                    national = Boolean.Parse(consoleValue);
                    break;
                } while (true);
            }
            else
            {
                filename = args[0];
                separator = args[1];
                exportType = args[2];
                national = Boolean.Parse(args[3]);
            }
            Debug.WriteLine($"Comando={filename} {separator} {exportType} {national}");
            
            var content = readFile(filename);
            if (content == null || content.Length <= 1)
            {
                Console.WriteLine("Il file non esiste o è vuoto");
                Console.ReadKey();
                return;
            }
            var contentDictionary = parseFile(content, separator);
            loadCategories();
            loadNations();
            switch(exportType.ToLower())
            {
                case "mispeaker":
                    MiSpeakerConverter(contentDictionary, national);
                    break;
                case "tvg":
                    TVGConverter(contentDictionary, national);
                    break;
            }
            Console.WriteLine("\nFile esportato. Premere un tasto per chiudere la finestra");
            Console.ReadKey();
        }
        private static void loadCategories()
        {
            categories.Clear();
            var lines = File.ReadAllLines("categorie.csv");
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
            var lines = File.ReadAllLines("nazioni.csv");
            var splitVal = new char[] { ';' };
            foreach(var line in lines)
            {
                var values = line.Split(splitVal, StringSplitOptions.None);
                nations.Add(values[1].Trim(), new Tuple<string, string>(values[0].Trim(), values[2].Trim()));
            }
        }
        private static string[] readFile(string path)
        {
            if(File.Exists(path))
                return File.ReadAllLines(path);
            return null;
        }
        private static List<Dictionary<string,string>> parseFile(string[] content, string separator)
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
            return listContent.OrderBy(x => x["Batteria"]).ThenBy(x=>x["Acqua"]).ToList();
        }
        private static void MiSpeakerConverter(List<Dictionary<string, string>> fields, bool isNational)
        {
            using (var file = new StreamWriter(new FileStream($@"C:\Users\pinoe\Desktop\PinoExport_MiSpeaker_{(isNational ? "Naz" : "Int")}.csv", FileMode.Create), Encoding.UTF8))
            {
                if(isNational)
                {

                }
                else
                    file.WriteLine($"Batteria;Acqua;Pettorale;Flag;Cognome;SecondoNome;Società;Atleta1;Atleta2;Atleta3;Atleta4");
                foreach(var row in fields)
                {
                    if (isNational)
                    {

                    }
                    else
                    {
                        var isTeam = !string.IsNullOrEmpty(row["Atleta3"].Trim()); //ci sono più di due atleti
                        var flag = $@"Flags3D\{(getFlagName(row["Nazione"].Trim()))}.png";
                        var teamName = getTeamName(row["Nazione"].Trim());
                        var surname = getSurname(row["Nazione"].Trim(), isTeam);
                        file.WriteLine($"{row["Batteria"]};{row["Acqua"]};{row["Pettorale"]};{flag};{surname};{(isTeam ? "" : row["Atleta1"].Replace("|", " "))};{teamName};{row["Atleta1"].Replace("|", " ")};{row["Atleta2"].Replace("|", " ")};{row["Atleta3"].Replace("|", " ")};{row["Atleta4"].Replace("|", " ")};");
                    }
                }
            }
        }
        private static string getFlagName(string nation)
        {
            var shortNation = getShortSurname(nation);
            if (nations.ContainsKey(nation))
                return nations[nation].Item1;
            else if (nations.ContainsKey(shortNation))
                return nations[shortNation].Item1;
            else
                Console.WriteLine($"Bandiera non trovata: {nation}");
            return nation.Substring(0,3);
        }
        private static string getTeamName(string nation)
        {
            var shortNation = getShortSurname(nation);
            if (nations.ContainsKey(nation))
                return nations[nation].Item2.ToUpper();
            else if (nations.ContainsKey(shortNation))
                return nations[shortNation].Item2.ToUpper();
            else
                Console.WriteLine($"Team - Nazione ({nation}) non trovata - Short ({nation})");
            return string.Empty;
        }
        private static string getSurname(string nation, bool isTeam)
        {
            var shortNation = getShortSurname(nation);
            if (nations.ContainsKey(nation) && isTeam)
                return nations[nation].Item2.ToUpper();
            else if(nations.ContainsKey(shortNation) && isTeam)
                return nations[shortNation].Item2.ToUpper();
            else if(isTeam)
                Console.WriteLine($"Surname - Nazione ({nation}) non trovata - Short ({shortNation})");
            return shortNation;
        }
        private static string getShortSurname(string nation)
        {
            var shortNation = nation.ToString();
            if (shortNation.Where(x => x == ' ').Count() >= 2)
            {
                var parts = shortNation.Split(new char[] { ' ' });
                shortNation = $"{parts[0]} {parts[1].Substring(0, 1)}{parts[2]}";
            }
            return shortNation;
        }
        private static void TVGConverter(List<Dictionary<string, string>> fields, bool isNational)
        {
            var groups = fields.GroupBy(x => x["Batteria"]);
            foreach (var group in groups)
            {
                Console.WriteLine($"Batteria {(group.ElementAt(0)["Batteria"])}: {group.Count()} concorrenti");
            }
            /*
                Acqua 
                Pettorale   
                Flag 
                Atleta  
                Societa 
                Societa1    
                Atleta1 
                Atleta2 
                Atleta3 
                Atleta4 
                Atleta5 
                Atleta6 
                Atleta7 
                Atleta8 
                Atleta9
                Categoria_Int
                Categoria_desc
                Batteria
                Acqua
                Equipaggio
            */

        }
        private static void credits()
        {
            Console.WriteLine("Software sviluppato da Giuseppe Elefante <giuseppe.elefante90@gmail.com>\nA.S.D. Cronometristi Salernitani \"Raffaele Marra\"\n\n");
        }
    }
    
}
