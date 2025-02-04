using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExcelExporter
{
    public class HojaExcel
    {
        public string Nombre { get; set; } = "";
        public string Titulo { get; set; } = "";
        public DataTable? Parametros { get; set; }
        public string SubTitulo { get; set; } = "";
        public DatosToExcelObject? Datos { get; set; }

        public HojaExcel()
        {
        }
    }
}
