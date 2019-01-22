using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;
using LinqToExcel.Attributes;

namespace ProjectScope.Client.Models
{
    public class UserStory
    {
        public string Name { get; set; }
        [ExcelColumn("ZEB nr")]
        public string Reference { get; set; }
        public string Scope { get; set; }
        [ExcelColumn("Additional comments")]
        public string Comments { get; set; }
        [ExcelColumn("Depends On ")]
        public string DependsOn { get; set; }
    }
}
