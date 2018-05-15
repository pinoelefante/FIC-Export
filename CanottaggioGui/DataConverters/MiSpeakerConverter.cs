using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CanottaggioGui.DataConverters
{
    public class MiSpeakerConverter : ConverterBasics
    {
        public override bool ConvertInternational(List<Dictionary<string, string>> fields, string title = "")
        {
            OutputStream.AppendLine("\nAvvio esportazione MiSpeaker Internazionale");
            var now = DateTime.Now;
            var filename = string.Format("Export_MiSpeaker_Int_{0:D2}{1:D2}{2:D4}.csv", now.Day, now.Month, now.Year);
            try
            {
                using (var file = new StreamWriter(new FileStream($@"{GetDesktopPath()}\{filename}", FileMode.Create), Encoding.UTF8))
                {
                    var buffer = new StringBuilder(16384); //16384 = 16KB
                    buffer.AppendLine($"Batteria;Acqua;Pettorale;Flag;Cognome;SecondoNome;Società;Atleta1;Atleta2;Atleta3;Atleta4");
                    foreach (var row in fields)
                    {
                        var isTeam = row.ContainsKey("Atleta3") && !string.IsNullOrEmpty(row["Atleta3"]); //ci sono più di due atleti
                        var flag = $@"Flags3D\{(GetFlagName(row["Nazione"].Trim()))}.png";
                        var teamName = GetTeamNameInt(row["Nazione"].Trim());
                        var surname = GetSurnameInt(row["Nazione"].Trim(), isTeam);
                        buffer.AppendLine(string.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}",
                            (row.ContainsKey("Batteria") ? row["Batteria"] : ""),
                            (row.ContainsKey("Acqua") ? row["Acqua"] : ""),
                            (row.ContainsKey("Pettorale") ? row["Pettorale"] : ""),
                            flag,
                            surname,
                            (!isTeam && row.ContainsKey("Atleta1") ? row["Atleta1"].Replace("|", " ") : ""),
                            teamName,
                            (row.ContainsKey("Atleta1") && !string.IsNullOrEmpty(row["Atleta1"]) ? row["Atleta1"].Replace('|', ' ') : ""),
                            (row.ContainsKey("Atleta2") && !string.IsNullOrEmpty(row["Atleta2"]) ? row["Atleta2"].Replace('|', ' ') : ""),
                            (row.ContainsKey("Atleta3") && !string.IsNullOrEmpty(row["Atleta3"]) ? row["Atleta3"].Replace('|', ' ') : ""),
                            (row.ContainsKey("Atleta4") && !string.IsNullOrEmpty(row["Atleta4"]) ? row["Atleta4"].Replace('|', ' ') : "")
                        ));
                    }
                    OutputStream.AppendLine($"Salvataggio file {filename} sul desktop");
                    file.Write(buffer.ToString());
                }
                OutputStream.AppendLine("Esportazione MiSpeaker completata con successo");
                return true;
            }
            catch (Exception e)
            {
                OutputStream.AppendLine($"Si e' verificato un errore durante l'esportazione di MiSpeaker\n{e.Message}\n{e.StackTrace}");
                return false;
            }
        }

        public override bool ConvertNational(List<Dictionary<string, string>> fields, string title = "")
        {
            OutputStream.AppendLine("\nAvvio esportazione MiSpeaker");
            var now = DateTime.Now;
            var filename = string.Format("Export_MiSpeaker_Naz_{0:D2}{1:D2}{2:D4}.csv", now.Day, now.Month, now.Year);
            try
            {
                using (var file = new StreamWriter(new FileStream($@"{GetDesktopPath()}\{filename}", FileMode.Create), Encoding.UTF8))
                {
                    var buffer = new StringBuilder(16384); //16384 = 16KB
                    OutputStream.AppendLine("ATTENZIONE: L'export nazionale potrebbe non essere corretto in quanto non testato");
                    buffer.AppendLine($"Batteria;Acqua;Pettorale;Atleta;Societa;Societa1;Atleta1;Atleta2;Atleta3;Atleta4;soc;Categoria[Categoria2];Descr_Cat");
                    foreach (var row in fields)
                    {
                        var isTeam = row.ContainsKey("Atleta3") && !string.IsNullOrEmpty(row["Atleta3"]); //ci sono più di due atleti
                        buffer.AppendLine($"{row["Batteria"]};{row["Acqua"]};{row["Pettorale"]};Atleta;Societa;Societa1;{row["Atleta1"].Replace("|", " ")};{row["Atleta2"].Replace("|", " ")};{row["Atleta3"].Replace("|", " ")};{row["Atleta4"].Replace("|", " ")};{row["Atleta5"].Replace("|", " ")};{row["Atleta6"].Replace("|", " ")};{row["Atleta7"].Replace("|", " ")};{row["Atleta8"].Replace("|", " ")};{row["Atleta9"].Replace("|", " ")};soc;{row["Categoria2"]};{GetCategoryDescription(row["Categoria2"], row["Categoria"])}");
                    }
                    OutputStream.AppendLine($"Salvataggio file {filename} sul desktop");
                    file.Write(buffer.ToString());
                }
                OutputStream.AppendLine("Esportazione MiSpeaker completata con successo");
                return true;
            }
            catch (Exception e)
            {
                OutputStream.AppendLine($"Si e' verificato un errore durante l'esportazione di MiSpeaker\n{e.Message}\n{e.StackTrace}");
                return false;
            }
        }
    }
}
