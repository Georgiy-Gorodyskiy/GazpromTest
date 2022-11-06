using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GazpromTest.Models
{
    public static class ServiceManager
    {
        public static ExcelService ExcelService { get; set; }
        public static void Init()
        {
            ExcelService = new ExcelService();
        }
    }
}
