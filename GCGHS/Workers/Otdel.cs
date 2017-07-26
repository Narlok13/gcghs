using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GCGHS.Workers
{
    public class Otdel
    {
        string otdelName;
        string otdelNumber;
        string okrug;

        public Otdel(string otdelNumb, string otdelName, string okrug)
        {
            OtdelNumber = otdelNumb;
            OtdelName = otdelName;
            Okrug = okrug;
        }

        public string OtdelName { get => otdelName; set => otdelName = value; }
        public string OtdelNumber { get => otdelNumber; set => otdelNumber = value; }
        public string Okrug { get => okrug; set => okrug = value; }
    }
}
