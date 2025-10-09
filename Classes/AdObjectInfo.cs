using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessManager.Classes
{
    public class AdObjectInfo
    {
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string Description { get; set; }
        public string Email { get; set; }
        public string Telephone { get; set; }
        public string EmployeeId { get; set; }
        public string DistinguishedName { get; set; }

        public override string ToString() => DisplayName ?? Name;

        public bool HasOwner { get; set; } = true;
    }
}
