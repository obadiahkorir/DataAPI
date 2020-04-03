using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CUEReceiver
{
    public class StudentsResearch
    {
        public int Code { get; set; }
        public String Campus { get; set; }
        public String Category { get; set; }
        public String Domain { get; set; }
        public String SubDomain { get; set; }
        public String PublicationType { get; set; }
        public String Title { get; set; }
        public String Description { get; set; }
        public String Link { get; set; }
        public String PatentingOrganisation { get; set; }
        public String PatentNo { get; set; }
        public int PatentYear { get; set; }
        public String AuthorIds { get; set; }
    }
}