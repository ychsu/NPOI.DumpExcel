using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.DumpExcel.Test.Models
{
    public class Parent
    {
        public IEnumerable<Foo> Foos { get; set; }

        public IEnumerable<Member> Members { get; set; }
    }
}
