using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CUEReceiver
{
    public class StudentJson
    {
        public String column_map { get; set; }
        public String userUserName { get; set; }
        public String userPassword { get; set; }
        public List<String[]> data { get; set; }
    }
}