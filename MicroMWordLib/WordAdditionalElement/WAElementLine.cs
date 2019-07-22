using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MicroMWordLib.WordAdditionalElement
{
    public class WAElementLine
    {
        private List<WAElementLineField> prp_Fields;

        public List<WAElementLineField> Fields { get => prp_Fields; set => prp_Fields = value; }

        public WAElementLine()
        {
            prp_Fields = new List<WAElementLineField>();
        }
    }
}
