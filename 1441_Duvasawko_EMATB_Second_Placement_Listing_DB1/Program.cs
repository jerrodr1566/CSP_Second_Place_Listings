using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _1441_Duvasawko_EMATB_Second_Placement_Listing_DB1
{
    class Program
    {
        public static bool TestMode = false;
        public static bool IsSQL = false;

        static void Main(string[] args)
        {
            _1441_Duvasawko_EMATB_Second_Placement_Listing_DB1 oDuvo = new _1441_Duvasawko_EMATB_Second_Placement_Listing_DB1(TestMode, IsSQL, "DB1");
        }
    }
}
