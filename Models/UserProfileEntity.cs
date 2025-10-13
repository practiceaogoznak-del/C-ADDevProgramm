using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessManager.Models
{
    public class UserProfileEntity
    {
        public int Id { get; set; }
        public string? Login { get; set; }
        public string? FullName { get; set; }
        public string? Description { get; set; }
        public bool Enabled { get; set; }
        public DateTime UpdatedAt { get; set; }
    }
}
