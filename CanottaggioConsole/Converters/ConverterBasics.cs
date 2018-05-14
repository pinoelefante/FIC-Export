using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CanottaggioConsole.Converters
{
    public abstract class ConverterBasics
    {
        public Dictionary<string, Tuple<string, string>> Categories { get; set; }
        public Dictionary<string, Tuple<string, string>> Nations { get; set; }
        public Dictionary<string, string> Teams { get; set; }

        public List<Dictionary<string, string>> Convert(List<Dictionary<string, string>> fields, bool isNational)
        {
            return isNational ? ConvertNational(fields) : ConvertInternational(fields);
        }
        public abstract List<Dictionary<string,string>> ConvertInternational(List<Dictionary<string, string>> fields);
        public abstract List<Dictionary<string, string>> ConvertNational(List<Dictionary<string, string>> fields);
    }
}
