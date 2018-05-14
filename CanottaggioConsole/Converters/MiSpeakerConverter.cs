using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CanottaggioConsole.Converters
{
    public class MiSpeakerConverter : ConverterBasics
    {
        public string CSVSeparator { get; set; }
        public override List<Dictionary<string, string>> Convert(List<Dictionary<string, string>> fields, bool isNational)
        {
            throw new NotImplementedException();
        }
    }
}
