using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PayStubCreator
{
    public class Company
    {
        public string Name { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string PostCode { get; set; }

        public string Header()
        {
            return string.Format("{0}\n{1}\n{2} {3}", Name, Address, PostCode, City);
        }
    }
}
