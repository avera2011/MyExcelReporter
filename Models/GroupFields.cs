using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SixLabors.ImageSharp;
namespace MyExcelExporter
{
    public class GroupFields
    {

        public int FromField { get; set; }
        public int FieldsQuantity { get; set; }
        public Color RGBBackColor { get; set; }
        public string Caption { get; set; }

        public Color RGBForeColor { get; set; }
        public GroupFields()
        {
            FromField = 1;
            FieldsQuantity = 2;
            RGBBackColor = Color.ParseHex("#05265A");
            RGBForeColor = Color.White;
            Caption = "";
        }
    }
}
