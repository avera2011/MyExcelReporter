
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using SixLabors.ImageSharp;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

using SixLabors.ImageSharp.Formats.Png;

namespace MyExcelExporter
{
    public class ReporteExcel
    {
        int row = 1;
        int col = 1;
        XLWorkbook workbook = new XLWorkbook();
        IXLWorksheet worksheet;
        List<DatosToExcelObject> _DATOS = new List<DatosToExcelObject>();
        public Image? Logo { get; set; }
        string nh = "";

        public ReporteExcel(Image? logo, string NombreHoja = "Hoja 1")
        {
            ;
            Logo = logo;
            nh = NombreHoja;
            worksheet = AgregarHoja(nh);
        }
        public IXLWorksheet AgregarHoja(string nhoja)
        {
            row = 1;
            col = 1;
            nh = nhoja;
            var w = workbook.Worksheets.Add(nh);
            if (Logo != null)
            {
                using var memoryStream = new MemoryStream();
                Logo.Save(memoryStream, new PngEncoder()); // Save as PNG
                memoryStream.Seek(0, SeekOrigin.Begin);

                // Add the image to the worksheet
                var pic = w.AddPicture(memoryStream, "MyImage" + workbook.Worksheets.Count);
                pic.MoveTo(0, 0);
                var altura = pic.Height;
                if (altura > 90)
                {
                    var percent = 90D / altura;
                    pic.Placement = ClosedXML.Excel.Drawings.XLPicturePlacement.FreeFloating;
                    pic.Scale(Math.Round(percent, 2), true);
                }
                w.Row(1).Height = 95;
            }
            worksheet = w;
            return worksheet;
            //workbook.Worksheets[1].Cell[1, 1].Copy(worksheet.Cell[1, 1]);

        }
        public void AgregarTablaSinFormato(DatosToExcelObject tabla)
        {
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
            int ini_row = row;
            int ini_col = col;
            //imprime agrupaciones de columnas
            if (tabla.Agrupaciones.Count > 0)
            {




                foreach (GroupFields grupo in tabla.Agrupaciones)
                {
                    var gru = worksheet.Range(row, ini_col + (grupo.FromField - 1), row, ini_col + (grupo.FromField - 1) + (grupo.FieldsQuantity - 1));
                    if (gru != null)
                    {
                        gru.Merge();
                        gru.Value = grupo.Caption;
                        gru.Style.Fill.PatternType = XLFillPatternValues.Solid;
                        gru.Style.Fill.BackgroundColor = XLColor.FromHtml(grupo.RGBBackColor.ToHex()); //grupo.RGBBackColor;
                        gru.Style.Font.FontColor = XLColor.FromHtml(grupo.RGBForeColor.ToHex());
                        gru.Style.Font.FontSize = float.Parse((tabla.DefaultForeSize + 1).ToString());
                        gru.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }




                }
                row++;
            }
            //IMPREIME LA TABLA
            if (tabla.Datos != null)
            {
                //CAMBIA EL CAPTIO DE LA COLUMNA POR SUS ALIAS , si los tiene
                foreach (DataColumn c in tabla.Datos.Columns)
                {
                    c.Caption = char.ToUpper(c.ColumnName[0]) + c.ColumnName.Substring(1);
                    if (tabla.Field_Alias.Count > 0)
                    {
                        for (int i = 0; i < tabla.Field_Alias.Count; i++)
                        {
                            if (tabla.Field_Alias[i] != "")
                            {
                                string[] fa = tabla.Field_Alias[i].Split('|');
                                if (fa.Length == 2)
                                {
                                    if (fa[0].ToLower() == c.ColumnName.ToLower())
                                    {
                                        c.Caption = fa[1];
                                        break;
                                    }
                                }
                            }
                        }
                    }

                    col++;
                }
                col = ini_col;
                //inserta la tabla
                var tablxl = worksheet.Cell(row, col).InsertTable(tabla.Datos, true);

                //obtiene la cabecera
                var hr = tablxl.HeadersRow();
                foreach (DataColumn c in tabla.Datos.Columns)
                {
                    if (tabla.markBorder)

                    {
                        hr.Cell(col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        hr.Cell(col).Style.Border.OutsideBorderColor = XLColor.FromHtml(tabla.DefaultBorderColor.ToHex());
                    }



                    hr.Cell(col).Style.Fill.PatternType = XLFillPatternValues.Solid;
                    hr.Cell(col).Style.Fill.BackgroundColor = XLColor.FromHtml(tabla.RGBHeaderBackColor.ToHex());
                    hr.Cell(col).Style.Font.FontColor = XLColor.FromHtml(tabla.RGBHeaderForeColor.ToHex());
                    hr.Cell(col).Style.Font.FontSize = float.Parse((tabla.DefaultForeSize + 1).ToString());
                    switch (c.DataType.Name.ToLower())
                    {
                        case "int16":
                        case "int32":
                        case "int64":
                        case "integer":
                        case "int":
                        case "decimal":

                        case "double":

                            hr.Cell(col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                            break;

                        case "datetime":
                            hr.Cell(col).Style.DateFormat.Format = "dd/MM/yyyy HH:mm:ss";
                            break;
                        default:
                            break;
                    }

                    col++;
                }





                //tablxl.ShowHeaderRow = false;
                row += tabla.Datos.Rows.Count + 1;
                col = ini_col;
                //IMPRIME FOOTER
                DataTable clon_row = tabla.Datos.Clone();
                clon_row.Rows.Add();
                var totales = new List<TotalGroup>();
                if (tabla.TotalFields.Count > 0)
                {

                    //SE CREA UNA LISTA DE TOTALIZACIÓN

                    if (!string.IsNullOrEmpty(tabla.TotalFieldGroup))
                    {
                        DataView view = new DataView(tabla.Datos);
                        DataTable distinctValues = view.ToTable(true, tabla.TotalFieldGroup);
                        for (int i = 0; i < distinctValues.Rows.Count; i++)
                        {
                            totales.Add(new TotalGroup() { Descripcion = distinctValues.Rows[i][tabla.TotalFieldGroup].ToString() ?? "0", value = 0 });

                        }
                        // totales.Add(new TotalGroup() { Descripcion = "", value = 0 });
                    }
                    totales.Add(new TotalGroup() { Descripcion = "", value = 0 });
                    //final de la tabla de datos


                    foreach (var tg in totales)
                    {
                        col = ini_col;

                        foreach (DataColumn c in tabla.Datos.Columns)
                        {
                            var global_index_total_field = 0;
                            decimal valor = 0;
                            bool estotal = false;
                            for (int index_total_field = 0; index_total_field < tabla.TotalFields.Count; index_total_field++)
                            {
                                global_index_total_field = index_total_field;
                                string tot = tabla.TotalFields[index_total_field].ToString();

                                string[] totar = tot.Split('|');
                                string nn = tot;
                                decimal cv = 0;
                                if (totar.Length > 1)
                                {
                                    cv = decimal.Parse(totar[1], NumberStyles.AllowExponent | NumberStyles.AllowLeadingSign | NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
                                    nn = totar[0];
                                }
                                if (c.ColumnName.ToLower() == nn.ToLower())
                                {
                                    estotal = true;
                                    if (totar.Length == 2)
                                    {
                                        valor = cv;
                                    }
                                    else
                                    {
                                        if (tg.Descripcion == "")
                                        {
                                            for (int i = 0; i < tabla.Datos.Rows.Count; i++)
                                            {
                                                if (tabla.Datos.Rows[i][c].ToString() != "")
                                                {
                                                    valor += decimal.Parse(tabla.Datos.Rows[i][c].ToString() ?? "0", NumberStyles.AllowExponent | NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            DataRow[] arr = tabla.Datos.Select(tabla.TotalFieldGroup + "='" + tg.Descripcion + "'");
                                            for (int i = 0; i < arr.Length; i++)
                                            {
                                                if (arr[i][c].ToString() != "")
                                                {
                                                    valor += decimal.Parse(arr[i][c].ToString() ?? "0", NumberStyles.AllowExponent | NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
                                                }
                                            }

                                        }

                                    }
                                    break;
                                }
                            }

                            if (estotal)
                            {
                                switch (c.DataType.Name.ToLower())
                                {
                                    case "int16":
                                    case "int32":
                                    case "integer":
                                    case "int":
                                    case "int64":
                                        worksheet.Cell(row, col).Style.NumberFormat.Format = "0";
                                        break;
                                    default:
                                        worksheet.Cell(row, col).Style.NumberFormat.Format = "0.00";
                                        break;
                                }
                                worksheet.Cell(row, col).Style.Font.FontSize = float.Parse(tabla.DefaultForeSize.ToString());
                                worksheet.Cell(row, col).Value = valor;

                                if (clon_row.Columns[c.ColumnName] != null)
                                {
                                    clon_row.Columns[c.ColumnName]!.ReadOnly = false;
                                }
                                clon_row.Rows[0][c.ColumnName] = valor;

                                worksheet.Cell(row, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                                worksheet.Cell(row, col).Style.NumberFormat.Format = "0.00";

                                worksheet.Cell(row, col).Style.Fill.PatternType = XLFillPatternValues.Solid;
                                worksheet.Cell(row, col).Style.Fill.BackgroundColor = XLColor.FromHtml(tabla.RGBFootBackColor.ToHex());
                                worksheet.Cell(row, col).Style.Font.FontColor = XLColor.FromHtml(tabla.RGBFootForeColor.ToHex());
                                worksheet.Cell(row, col).Style.Font.Bold = true;
                                if (tabla.markBorder)
                                {
                                    worksheet.Cell(row, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                    worksheet.Cell(row, col).Style.Border.OutsideBorderColor = XLColor.FromHtml(tabla.DefaultBorderColor.ToHex());
                                }


                                if (col - 1 > 0 && global_index_total_field == 0)
                                {
                                    worksheet.Cell(row, col - 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                                    worksheet.Cell(row, col - 1).Style.Border.OutsideBorderColor = XLColor.FromHtml(tabla.DefaultBorderColor.ToHex());
                                    worksheet.Cell(row, col - 1).Style.Fill.PatternType = XLFillPatternValues.Solid;
                                    worksheet.Cell(row, col - 1).Style.Fill.BackgroundColor = XLColor.FromHtml(tabla.RGBFootBackColor.ToHex());
                                    worksheet.Cell(row, col - 1).Style.Font.FontColor = XLColor.FromHtml(tabla.RGBFootForeColor.ToHex());
                                    worksheet.Cell(row, col - 1).Style.Fill.BackgroundColor = XLColor.FromHtml(tabla.RGBFootBackColor.ToHex());
                                    worksheet.Cell(row, col - 1).Style.Font.Bold = true;
                                    worksheet.Cell(row, col - 1).Value = tg.Descripcion;
                                }

                            }


                            col++;
                        }
                        row++;
                    }

                    //tabla.Datos.Rows.Add(clon_row.Rows[0].ItemArray);
                }

                _DATOS.Add(new DatosToExcelObject());
                if (_DATOS[_DATOS.Count - 1]!.Datos != null)
                {
                    _DATOS[_DATOS.Count - 1]!.Datos = tabla.Datos.Copy();
                    _DATOS[_DATOS.Count - 1]!.Datos?.Rows.Add(clon_row.Rows[0].ItemArray);
                }
            }



            //_DATOS.Add(new DatosToExcelObject());
            //_DATOS[_DATOS.Count - 1].Datos = tabla.Datos.Copy();
            //_DATOS[_DATOS.Count - 1].Datos.Rows.Add(clon_row.Rows[0].ItemArray);
            //clon_row.Reset();


            //worksheet.Columns().AdjustToContents();
            col = ini_col;
            //row ++;
        }
        public void AgregarTabla(DatosToExcelObject tabla)
        {

            //IMPREIME CABECERA
            AgregarTablaSinFormato(tabla);

        }
        public void AgregarEspacio(double ancho = 50)
        {
            worksheet.Column(col).Width = ancho;

            col++;

        }
        public void Bajar(int rows = 1)
        {
            col = 1;
            row = row + rows;

        }
        public void Escribir(string texto, int fila = 0, int columna = 0, int mergecells = 1, int alineacion = 0, float FontSize = 9, bool negrilla = false
            , bool format = false, Color border_color = default, Color back_color = default, Color fore_color = default
            )
        {
            if (fila == 0) fila = row;
            if (columna == 0) columna = col;
            switch (alineacion)
            {
                case 1:
                    worksheet.Cell(fila, columna).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    break;
                case 2:
                    worksheet.Cell(fila, columna).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    break;
                default:
                    worksheet.Cell(fila, columna).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    break;
            }
            if (mergecells > 1)
            {
                worksheet.Range(fila, columna, fila, columna + mergecells - 1).Merge();
            }

            worksheet.Cell(fila, columna).Style.Font.FontSize = FontSize;
            worksheet.Cell(fila, columna).Style.Font.Bold = negrilla;
            worksheet.Cell(fila, columna).Value = texto;
            if (format)
            {
                //worksheet.Cell(fila, columna).Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, border_color);
                worksheet.Cell(fila, columna).Style.Fill.PatternType = XLFillPatternValues.Solid;
                worksheet.Cell(fila, columna).Style.Fill.BackgroundColor = XLColor.FromHtml(back_color.ToHex());
                worksheet.Cell(fila, columna).Style.Font.FontColor = XLColor.FromHtml(fore_color.ToHex());
            }


            col = columna + mergecells;
            row = fila;
        }
        public byte[] GetBytes()
        {
            //p.SaveAs(new System.IO.FileInfo("c:/archivoexcel.xlsx"));

            using var memoryStream = new MemoryStream();
            workbook.SaveAs(memoryStream);
            byte[] excelBytes = memoryStream.ToArray();
            return excelBytes;
        }
        public List<DatosToExcelObject> GetData()
        {
            return _DATOS;
        }


    }

}
