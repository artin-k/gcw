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
        public int Position { get; set; }

        public Bookmark(string name, int position)
        {
            Name = name;
            Position = position;
        }
    }


}
