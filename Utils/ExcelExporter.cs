using ClosedXML.Excel;
using IOLPortfolio.Models;

namespace IOLPortfolio.Utils
{
    public static class ExcelExporter
    {
        public static void Save(string filePath, List<Portfolio> items)
        {
            using var wb = new XLWorkbook();


            var ws = wb.AddWorksheet("Detalle");
            var headers = new[]
            {
                "Simbolo","Descripcion","Tipo","Mercado","Moneda",
                "Cantidad","UltimoPrecio","Ppc","Valorizado",
                "GananciaDinero","GananciaPorcentaje","VariacionDiaria"
            };
            for (int c = 0; c < headers.Length; c++)
                ws.Cell(1, c + 1).Value = headers[c];

            for (int i = 0; i < items.Count; i++)
            {
                var r = i + 2;
                var p = items[i];
                ws.Cell(r, 1).Value = p.Simbolo;
                ws.Cell(r, 2).Value = p.Descripcion;
                ws.Cell(r, 3).Value = p.Tipo;
                ws.Cell(r, 4).Value = p.Mercado;
                ws.Cell(r, 5).Value = p.Moneda;
                ws.Cell(r, 6).Value = p.Cantidad;
                ws.Cell(r, 7).Value = p.UltimoPrecio;
                ws.Cell(r, 8).Value = p.Ppc;
                ws.Cell(r, 9).Value = p.Valorizado;
                ws.Cell(r, 10).Value = p.GananciaDinero;
                ws.Cell(r, 11).Value = p.GananciaPorcentaje / 100m; 
                ws.Cell(r, 12).Value = p.VariacionDiaria / 100m;    
            }

            var used = ws.RangeUsed();
            used.CreateTable();
            ws.Columns().AdjustToContents();
            ws.Column(11).Style.NumberFormat.Format = "0.00%";
            ws.Column(12).Style.NumberFormat.Format = "0.00%";
            ws.Column(6).Style.NumberFormat.Format = "#,##0.####";
            ws.Column(7).Style.NumberFormat.Format = "#,##0.00";
            ws.Column(8).Style.NumberFormat.Format = "#,##0.00";
            ws.Column(9).Style.NumberFormat.Format = "#,##0.00";
            ws.Column(10).Style.NumberFormat.Format = "#,##0.00";

            var wsTipo = wb.AddWorksheet("Resumen por tipo");
            var resumen = items
                .GroupBy(p => p.Tipo ?? "")
                .Select(g => new { Tipo = g.Key, Valorizado = g.Sum(x => x.Valorizado) })
                .OrderByDescending(x => x.Valorizado)
                .ToList();

            wsTipo.Cell(1, 1).Value = "Tipo";
            wsTipo.Cell(1, 2).Value = "Valorizado";
            for (int i = 0; i < resumen.Count; i++)
            {
                wsTipo.Cell(i + 2, 1).Value = resumen[i].Tipo;
                wsTipo.Cell(i + 2, 2).Value = resumen[i].Valorizado;
            }
            wsTipo.RangeUsed().CreateTable();
            wsTipo.Columns().AdjustToContents();
            wsTipo.Column(2).Style.NumberFormat.Format = "#,##0.00";

            var wsTop = wb.AddWorksheet("Top 10 por valuacion");
            var top = items.OrderByDescending(p => p.Valorizado).Take(10).ToList();
            wsTop.Cell(1, 1).Value = "Simbolo";
            wsTop.Cell(1, 2).Value = "Descripcion";
            wsTop.Cell(1, 3).Value = "Valorizado";

            for (int i = 0; i < top.Count; i++)
            {
                wsTop.Cell(i + 2, 1).Value = top[i].Simbolo;
                wsTop.Cell(i + 2, 2).Value = top[i].Descripcion;
                wsTop.Cell(i + 2, 3).Value = top[i].Valorizado;
            }
            wsTop.RangeUsed().CreateTable();
            wsTop.Columns().AdjustToContents();
            wsTop.Column(3).Style.NumberFormat.Format = "#,##0.00";

            var wsTot = wb.AddWorksheet("Totales");
            wsTot.Cell(1, 1).Value = "Total Valorizado";
            wsTot.Cell(1, 2).Value = items.Sum(p => p.Valorizado);
            wsTot.Column(2).Style.NumberFormat.Format = "#,##0.00";
            wsTot.Columns().AdjustToContents();

            wb.SaveAs(filePath);
        }
    }
}
