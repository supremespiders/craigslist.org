using System.Collections.Generic;

namespace craigslist.org.Models
{
    public class Category
    {
        public string CategoryName { get; set; }
        public List<RealEstate> RealEstate { get; set; } = new List<RealEstate>();
    }
}