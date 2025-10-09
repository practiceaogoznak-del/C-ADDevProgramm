using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessManager.Classes
{
    public class ResourceRequestItem
    {
        public AdObjectInfo Resource { get; set; }
        public bool IsRequested { get; set; }
        public bool CurrentHasAccess { get; set; }


    }
}
