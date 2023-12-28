using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DrawDocument
{
    public class Sentence
    {
        public string Text { get; set; }
        public string Color { get; set; }
    }

    public class DrawParams
    {
        public string InputPath { get; set; }
        public List<Sentence> Sentences { get; set; }
        public string OutputPath { get; set; }
    }
}
