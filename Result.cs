using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SJResults
{
    class Result : Competitor
    {
        public string Place { get; set; }
        public string FirstLength { get; set; }
        public string SecondLength { get; set; }
        public string ThirdLength { get; set; }
        public string Point { get; set; }
        public string FirstNote { get; set; }
        public string SecondNote { get; set; }
    }
}
