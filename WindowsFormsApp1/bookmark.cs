using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class Bookmark
    {
        public string Name { get; set; }
        public int Start { get; set; }
        public int Length { get; set; }

        public Bookmark(string name, int start, int length = 0)
        {
            Name = name;
            Start = start;
            Length = length;
        }
    }

}
