using System;
using System.Collections.Generic;
using System.Data;
using SixLabors.ImageSharp;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace MyExcelExporter
{
    public class DatosToExcelObject
    {
        public DataTable? Datos { get; set; }
        public List<string> Field_Alias { get; set; }
        public string TotalFieldGroup { get; set; } = "";
        public List<string> TotalFields { get; set; }
        public bool markBorder { get; set; }
        public Color RGBHeaderBackColor { get; set; }
        public Color RGBHeaderForeColor { get; set; }
        public Color RGBFootBackColor { get; set; }
        public Color RGBFootForeColor { get; set; }
        public Color RGBAlterBackColor { get; set; }
        public Color RGBForeColor { get; set; }
        public double DefaultForeSize { get; set; }
        public Color DefaultBorderColor { get; set; }

        public double DefaultSpaceColumnWidth { get; set; }
        public double DefaultRowHeight { get; set; }
        public double DefaultRowHeaderHeight { get; set; }
        public double DefaultRowFooterHeight { get; set; }
        public List<GroupFields> Agrupaciones { get; set; }
        public DatosToExcelObject()
        {
            RGBHeaderBackColor = Color.ParseHex("#5F7FB1");
            RGBAlterBackColor = Color.ParseHex("#C7D1E1");
            RGBFootBackColor = Color.ParseHex("#5F7FB1");
            RGBHeaderForeColor = Color.White;
            RGBForeColor = Color.Black;
            RGBFootForeColor = Color.White;
            DefaultForeSize = 12;
            DefaultBorderColor = Color.Black;
            markBorder = false;
            DefaultSpaceColumnWidth = 25;
            DefaultRowHeight = 55;
            DefaultRowHeaderHeight = 50;
            DefaultRowFooterHeight = 55;
            Agrupaciones = new List<GroupFields>();
            TotalFields = new List<string>();
            Field_Alias = new List<string>();
        }


    }
}
