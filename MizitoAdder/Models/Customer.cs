using System;
using System.Collections.Generic;
using System.Text;

namespace MizitoAdder.Models
{
    public class Customer
    {
        public required int RowId { get; set; }
        public required string Name { get; set; }
        public string? ShopName { get; set; }
        public string? Telephone { get; set; }
        public string? Phone { get; set; }
        public string? Email { get; set; }
        public required string Address { get; set; } = "";
        public string? Tags { get; set; }
        public string? RepresentativeName { get; set; }
        public string? RepresentativePhone { get; set; }
    }
}
