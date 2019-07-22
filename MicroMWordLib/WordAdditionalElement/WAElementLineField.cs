using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MicroMWordLib.WordText;

namespace MicroMWordLib.WordAdditionalElement
{
    public class WAElementLineField
    {
        private List<WTextPart> prp_Elements;

        public List<WTextPart> Elements { get => prp_Elements; set => prp_Elements = value; }

        public WAElementLineField()
        {
            prp_Elements = new List<WTextPart>();
        }
    }
}
