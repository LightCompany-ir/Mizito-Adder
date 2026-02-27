using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace MizitoAdder.Models
{
    public class ImportDTO
    {
        public required int ImportId { get; set; }
        [DisplayName("نام مشتری")]
        public required string CustomerName { get; set; }
        [DisplayName("نام تجاری مشتری")]
        public string? ShopName { get; set; }
        [DisplayName("تلفن ثابت")]
        public string? Telephone { get; set; }
        [DisplayName("شماره همراه")]
        public required string Phone { get; set; }
        [DisplayName("آدرس")]
        public string? Address { get; set; }
        [DisplayName("توضیحات")]
        public string? Info { get; set; }
        [DisplayName("آدرس سایت")]
        public string? Website { get; set; }
        [DisplayName("آدرس ایمیل")]
        public string? Email { get; set; }
        [DisplayName("کد پستی")]
        public string? PostalCode { get; set; }
        [DisplayName("فکس")]
        public string? FaxNumber { get; set; }
        [DisplayName("کد اقتصادی")]
        public string? EconomicCode { get; set; }
        [DisplayName("شناسه ملی")]
        public string? NationalID { get; set; }
        [DisplayName("برچسب ها")]
        public string? Tags { get; set; }
        [DisplayName("نام نماینده حقیقی مشتری")]
        public string? RepresentativeName { get; set; }
        [DisplayName("سمت")]
        public string? RepresentativePosition { get; set; }
        [DisplayName("آدرس ایمیل نماینده")]
        public string? RepresentativeEmail { get; set; }
        [DisplayName("شماره موبایل نماینده")]
        public string? RepresentativePhone { get; set; }
        [DisplayName("تلفن ثابت نماینده")]
        public string? RepresentativeTelephone { get; set; }
    }
}
