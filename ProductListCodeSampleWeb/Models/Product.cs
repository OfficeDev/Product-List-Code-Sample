using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SharePointAppSampleWeb.Models
{
    public class Product
    {
        public int Id { get; set; }
        public string Title { get; set; }

        public string Description { get; set; }

        public string Price { get; set; }
    }
}