using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessManager.Models
{
    public class UserProfileModel
    {
        public string? Login { get; set; }
        public string? FullName { get; set; }
        public string? Description { get; set; }
        public bool Enabled { get; set; }
    }
}
