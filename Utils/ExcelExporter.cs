using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Chart;
using IOLPortfolio.Models;
using System.Drawing;

namespace IOLPortfolio.Utils
{
    public static class ExcelExporter
    {
        // Paleta: (fondo, texto, acento) por tipo normalizado
        private static readonly Dictionary<string, (Color Bg, Color Fg, Color Accent)> Paleta =
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["CEDEAR"]           = (Hex("BDD7EE"), Hex("1F4E79"), Hex("2E75B6")),
                ["Acciones"]         = (Hex("E2EFDA"), Hex("1E4620"), Hex("70AD47")),
                ["Títulos Públicos"] = (Hex("FCE4D6"), Hex("843C0C"), Hex("ED7D31")),
                ["Letras"]           = (Hex("FFF2CC"), Hex("7F6000"), Hex("FFC000")),
                ["FCI"]              = (Hex("E8D5F5"), Hex("4B1076"), Hex("9B59B6")),
                ["ON"]               = (Hex("D5F5E3"), Hex("1A6632"), Hex("27AE60")),
            };

        private static (Color Bg, Color Fg, Color Accent) Colores(string tipo) =>
            Paleta.TryGetValue(tipo, out var c) ? c : (Hex("F2F2F2"), Hex("333333"), Hex("595959"));

        public static void Save(string filePath, List<Portfolio> items)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);

            using var pkg = new ExcelPackage();

            BuildDashboard(pkg, items);
            BuildDetalle(pkg, items);
            BuildPnL(pkg, items);
            BuildTop10(pkg, items);

            // Dashboard como primera hoja activa
            pkg.Workbook.Worksheets["Dashboard"].Select();

            pkg.SaveAs(new FileInfo(filePath));
        }

        // ═══════════════════════════════════════════════════
        //  HOJA 1 — DASHBOARD
        // ═══════════════════════════════════════════════════
        private static void BuildDashboard(ExcelPackage pkg, List<Portfolio> items)
        {
            var ws = pkg.Workbook.Worksheets.Add("Dashboard");
            ws.View.ShowGridLines = false;
            ws.TabColor = Hex("2E75B6");

            // ── Título ──
            Merge(ws, "B2:J2", $"PORTAFOLIO IOL  ·  RESUMEN EJECUTIVO",
                18, bold: true, fg: Hex("1F4E79"), bg: Color.Transparent, halign: ExcelHorizontalAlignment.Left);
            Merge(ws, "B3:J3", $"Generado el {DateTime.Now:dd/MM/yyyy  HH:mm}",
                10, fg: Hex("595959"), halign: ExcelHorizontalAlignment.Left);
            ws.Row(2).Height = 30;

            // ── KPIs ──
            decimal totalVal = items.Sum(p => p.Valorizado);
            decimal totalGan = items.Sum(p => p.GananciaDinero);
            decimal costo = totalVal - totalGan;
            decimal pctGan = costo > 0 ? totalGan / costo * 100m : 0;
            decimal varDia = items.Sum(p => p.Valorizado * p.VariacionDiaria / 100m);

            var kpis = new[]
            {
                ("Total Valorizado",   $"$ {totalVal:N2}",          Hex("2E75B6")),
                ("Ganancia Acum. $",   $"{(totalGan>=0?"+":"")}{totalGan:N2}", totalGan  >= 0 ? Hex("27AE60") : Hex("C0392B")),
                ("Ganancia Acum. %",   $"{(pctGan >=0?"+":"")}{pctGan:N2}%",  pctGan    >= 0 ? Hex("27AE60") : Hex("C0392B")),
                ("Variación del Día $",$"{(varDia >=0?"+":"")}{varDia:N2}",   varDia    >= 0 ? Hex("27AE60") : Hex("C0392B")),
            };

            int kRow = 5;
            for (int k = 0; k < kpis.Length; k++)
            {
                int col = 2 + k * 2; // B, D, F, H
                var (label, value, color) = kpis[k];

                // Label band
                var lbl = ws.Cells[kRow, col, kRow, col + 1];
                lbl.Merge = true; lbl.Value = label;
                lbl.Style.Fill.PatternType = ExcelFillStyle.Solid;
                lbl.Style.Fill.BackgroundColor.SetColor(color);
                lbl.Style.Font.Color.SetColor(Color.White);
                lbl.Style.Font.Bold = true; lbl.Style.Font.Size = 9;
                lbl.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                lbl.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Row(kRow).Height = 20;

                // Value band
                var val = ws.Cells[kRow + 1, col, kRow + 1, col + 1];
                val.Merge = true; val.Value = value;
                val.Style.Fill.PatternType = ExcelFillStyle.Solid;
                val.Style.Fill.BackgroundColor.SetColor(Hex("F8F9FA"));
                val.Style.Font.Size = 13; val.Style.Font.Bold = true;
                val.Style.Font.Color.SetColor(color);
                val.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                val.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                val.Style.Border.BorderAround(ExcelBorderStyle.Medium, color);
                ws.Row(kRow + 1).Height = 28;
            }

            // ── Tabla distribución por tipo (datos para el pie chart) ──
            int tRow = 9;
            var resumen = items
                .GroupBy(p => NormTipo(p.Tipo))
                .Select(g => new {
                    Tipo = g.Key,
                    Valorizado = g.Sum(x => x.Valorizado),
                    Ganancia = g.Sum(x => x.GananciaDinero),
                    Cant = g.Count()
                })
                .OrderByDescending(x => x.Valorizado)
                .ToList();

            foreach (var (txt, c) in new[] {
                ("Tipo de Activo", 2), ("Valorizado $", 3), ("% del Total", 4),
                ("Ganancia $", 5), ("Activos", 6) })
            {
                SetHeader(ws.Cells[tRow, c], txt);
            }
            ws.Row(tRow).Height = 22;

            for (int i = 0; i < resumen.Count; i++)
            {
                int row = tRow + 1 + i;
                var r = resumen[i];
                decimal pct = totalVal > 0 ? r.Valorizado / totalVal : 0;
                var (bg, fg, _) = Colores(r.Tipo);
                Color rowBg = i % 2 == 0 ? bg : Blend(bg, Color.White, 0.35f);

                ws.Cells[row, 2].Value = r.Tipo;
                ws.Cells[row, 3].Value = r.Valorizado;  ws.Cells[row, 3].Style.Numberformat.Format = "#,##0.00";
                ws.Cells[row, 4].Value = pct;           ws.Cells[row, 4].Style.Numberformat.Format = "0.00%";
                ws.Cells[row, 5].Value = r.Ganancia;    ws.Cells[row, 5].Style.Numberformat.Format = "#,##0.00";
                ws.Cells[row, 6].Value = r.Cant;

                var rng = ws.Cells[row, 2, row, 6];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(rowBg);
                rng.Style.Font.Color.SetColor(fg);
                rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                rng.Style.Border.Bottom.Color.SetColor(Hex("D9D9D9"));
                ws.Row(row).Height = 18;

                ws.Cells[row, 5].Style.Font.Color.SetColor(r.Ganancia >= 0 ? Hex("375623") : Hex("9C0006"));
                ws.Cells[row, 5].Style.Font.Bold = true;
            }

            ws.Column(1).Width = 2; ws.Column(2).Width = 22; ws.Column(3).Width = 17;
            ws.Column(4).Width = 13; ws.Column(5).Width = 17; ws.Column(6).Width = 10;
            ws.Column(7).Width = 2;

            // ── Pie Chart ──
            int dataStart = tRow + 1;
            int dataEnd   = tRow + resumen.Count;

            var pie = (ExcelPieChart)ws.Drawings.AddChart("PieDistribucion", eChartType.Pie3D);
            pie.Series.Add(ws.Cells[dataStart, 3, dataEnd, 3], ws.Cells[dataStart, 2, dataEnd, 2]);
            pie.Title.Text = "Distribución por Tipo";
            pie.Title.Font.Size = 12;
            pie.Legend.Position = eLegendPosition.Bottom;
            pie.DataLabel.ShowPercent = true;
            pie.DataLabel.ShowCategory = false;
            pie.SetPosition(4, 0, 8, 0);   // fila 5, col I
            pie.SetSize(380, 300);
        }

        // ═══════════════════════════════════════════════════
        //  HOJA 2 — DETALLE
        // ═══════════════════════════════════════════════════
        private static void BuildDetalle(ExcelPackage pkg, List<Portfolio> items)
        {
            var ws = pkg.Workbook.Worksheets.Add("Detalle");
            ws.View.ShowGridLines = false;
            ws.View.FreezePanes(3, 1);
            ws.TabColor = Hex("595959");

            // Título
            Merge(ws, "A1:L1", $"Detalle del Portafolio  —  {DateTime.Now:dd/MM/yyyy}",
                13, bold: true, fg: Hex("1F4E79"), bg: Hex("DEEAF1"), halign: ExcelHorizontalAlignment.Center);
            ws.Row(1).Height = 26;

            // Headers
            var headers = new[] {
                "Símbolo","Descripción","Tipo","Mercado","Moneda",
                "Cantidad","Últ. Precio","Precio Prom.","Valorizado",
                "Ganancia $","Ganancia %","Var. Diaria"
            };
            for (int c = 0; c < headers.Length; c++)
                SetHeader(ws.Cells[2, c + 1], headers[c]);
            ws.Row(2).Height = 22;

            // Filas de datos
            for (int i = 0; i < items.Count; i++)
            {
                int row = i + 3;
                var p = items[i];
                var tipo = NormTipo(p.Tipo);
                var (bg, fg, _) = Colores(tipo);
                Color rowBg = i % 2 == 0 ? bg : Blend(bg, Color.White, 0.4f);

                ws.Cells[row, 1].Value  = p.Simbolo;
                ws.Cells[row, 2].Value  = p.Descripcion;
                ws.Cells[row, 3].Value  = tipo;
                ws.Cells[row, 4].Value  = p.Mercado;
                ws.Cells[row, 5].Value  = p.Moneda;
                ws.Cells[row, 6].Value  = p.Cantidad;
                ws.Cells[row, 7].Value  = p.UltimoPrecio;
                ws.Cells[row, 8].Value  = p.Ppc;
                ws.Cells[row, 9].Value  = p.Valorizado;
                ws.Cells[row, 10].Value = p.GananciaDinero;
                ws.Cells[row, 11].Value = p.GananciaPorcentaje / 100m;
                ws.Cells[row, 12].Value = p.VariacionDiaria / 100m;

                var rng = ws.Cells[row, 1, row, 12];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(rowBg);
                rng.Style.Font.Color.SetColor(fg);
                rng.Style.Font.Size = 9;
                rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                rng.Style.Border.Bottom.Color.SetColor(Hex("D9D9D9"));
                ws.Row(row).Height = 16;

                ws.Cells[row, 6].Style.Numberformat.Format  = "#,##0.####";
                ws.Cells[row, 7].Style.Numberformat.Format  = "#,##0.00";
                ws.Cells[row, 8].Style.Numberformat.Format  = "#,##0.00";
                ws.Cells[row, 9].Style.Numberformat.Format  = "#,##0.00";
                ws.Cells[row, 10].Style.Numberformat.Format = "#,##0.00";
                ws.Cells[row, 11].Style.Numberformat.Format = "0.00%";
                ws.Cells[row, 12].Style.Numberformat.Format = "0.00%";

                // Semáforos
                SemaforoBold(ws.Cells[row, 10], p.GananciaDinero);
                Semaforo(ws.Cells[row, 11], p.GananciaPorcentaje);
                Semaforo(ws.Cells[row, 12], p.VariacionDiaria);
            }

            ws.Column(1).Width = 12;  ws.Column(2).Width = 36; ws.Column(3).Width = 18;
            ws.Column(4).Width = 12;  ws.Column(5).Width = 10; ws.Column(6).Width = 12;
            ws.Column(7).Width = 14;  ws.Column(8).Width = 14; ws.Column(9).Width = 16;
            ws.Column(10).Width = 16; ws.Column(11).Width = 12; ws.Column(12).Width = 12;
        }

        // ═══════════════════════════════════════════════════
        //  HOJA 3 — P&L POR TIPO
        // ═══════════════════════════════════════════════════
        private static void BuildPnL(ExcelPackage pkg, List<Portfolio> items)
        {
            var ws = pkg.Workbook.Worksheets.Add("P&L por Tipo");
            ws.View.ShowGridLines = false;
            ws.TabColor = Hex("70AD47");

            Merge(ws, "A1:G1", "Resumen P&L por Tipo de Activo",
                13, bold: true, fg: Hex("1F4E79"), bg: Hex("DEEAF1"), halign: ExcelHorizontalAlignment.Center);
            ws.Row(1).Height = 26;

            var hdrs = new[] { "Tipo","Activos","Valorizado $","% del Total","Ganancia $","Ganancia %","Var. Día $" };
            for (int c = 0; c < hdrs.Length; c++)
                SetHeader(ws.Cells[2, c + 1], hdrs[c]);
            ws.Row(2).Height = 22;

            decimal totalVal = items.Sum(p => p.Valorizado);

            var resumen = items
                .GroupBy(p => NormTipo(p.Tipo))
                .Select(g => new {
                    Tipo      = g.Key,
                    Cant      = g.Count(),
                    Valorizado= g.Sum(x => x.Valorizado),
                    Ganancia  = g.Sum(x => x.GananciaDinero),
                    VarDia    = g.Sum(x => x.Valorizado * x.VariacionDiaria / 100m)
                })
                .OrderByDescending(x => x.Valorizado)
                .ToList();

            for (int i = 0; i < resumen.Count; i++)
            {
                int row = i + 3;
                var r = resumen[i];
                var (bg, fg, _) = Colores(r.Tipo);
                Color rowBg = i % 2 == 0 ? bg : Blend(bg, Color.White, 0.4f);
                decimal pct  = totalVal > 0 ? r.Valorizado / totalVal : 0;
                decimal gPct = (r.Valorizado - r.Ganancia) > 0 ? r.Ganancia / (r.Valorizado - r.Ganancia) : 0;

                ws.Cells[row, 1].Value = r.Tipo;
                ws.Cells[row, 2].Value = r.Cant;
                ws.Cells[row, 3].Value = r.Valorizado;  ws.Cells[row, 3].Style.Numberformat.Format = "#,##0.00";
                ws.Cells[row, 4].Value = pct;           ws.Cells[row, 4].Style.Numberformat.Format = "0.00%";
                ws.Cells[row, 5].Value = r.Ganancia;    ws.Cells[row, 5].Style.Numberformat.Format = "#,##0.00";
                ws.Cells[row, 6].Value = gPct;          ws.Cells[row, 6].Style.Numberformat.Format = "0.00%";
                ws.Cells[row, 7].Value = r.VarDia;      ws.Cells[row, 7].Style.Numberformat.Format = "#,##0.00";

                var rng = ws.Cells[row, 1, row, 7];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(rowBg);
                rng.Style.Font.Color.SetColor(fg);
                rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                rng.Style.Border.Bottom.Color.SetColor(Hex("D9D9D9"));
                ws.Row(row).Height = 18;

                SemaforoBold(ws.Cells[row, 5], r.Ganancia);
                Semaforo(ws.Cells[row, 6], gPct);
                Semaforo(ws.Cells[row, 7], r.VarDia);
            }

            // Fila de totales
            int totRow = resumen.Count + 3;
            ws.Cells[totRow, 1].Value = "TOTAL";
            ws.Cells[totRow, 2].Value = items.Count;
            ws.Cells[totRow, 3].Value = totalVal;       ws.Cells[totRow, 3].Style.Numberformat.Format = "#,##0.00";
            ws.Cells[totRow, 4].Value = 1m;             ws.Cells[totRow, 4].Style.Numberformat.Format = "0.00%";
            ws.Cells[totRow, 5].Value = items.Sum(p => p.GananciaDinero);
                                                        ws.Cells[totRow, 5].Style.Numberformat.Format = "#,##0.00";
            ws.Cells[totRow, 7].Value = items.Sum(p => p.Valorizado * p.VariacionDiaria / 100m);
                                                        ws.Cells[totRow, 7].Style.Numberformat.Format = "#,##0.00";

            var totRng = ws.Cells[totRow, 1, totRow, 7];
            totRng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            totRng.Style.Fill.BackgroundColor.SetColor(Hex("1F4E79"));
            totRng.Style.Font.Color.SetColor(Color.White);
            totRng.Style.Font.Bold = true;
            totRng.Style.Border.Top.Style = ExcelBorderStyle.Medium;
            totRng.Style.Border.Top.Color.SetColor(Hex("16375E"));
            ws.Row(totRow).Height = 20;

            ws.Column(1).Width = 22; ws.Column(2).Width = 10; ws.Column(3).Width = 17;
            ws.Column(4).Width = 12; ws.Column(5).Width = 17; ws.Column(6).Width = 12;
            ws.Column(7).Width = 17;

            // Gráfico de columnas — Ganancia por tipo
            int ds = 3, de = resumen.Count + 2;
            var bar = (ExcelBarChart)ws.Drawings.AddChart("GananciaTipo", eChartType.ColumnClustered);
            bar.Series.Add(ws.Cells[ds, 5, de, 5], ws.Cells[ds, 1, de, 1]);
            bar.Series[0].Header = "Ganancia $";
            bar.Title.Text = "Ganancia por Tipo de Activo";
            bar.Title.Font.Size = 11;
            bar.Legend.Remove();
            bar.SetPosition(2, 0, 8, 0);
            bar.SetSize(400, 260);
        }

        // ═══════════════════════════════════════════════════
        //  HOJA 4 — TOP 10
        // ═══════════════════════════════════════════════════
        private static void BuildTop10(ExcelPackage pkg, List<Portfolio> items)
        {
            var ws = pkg.Workbook.Worksheets.Add("Top 10");
            ws.View.ShowGridLines = false;
            ws.TabColor = Hex("ED7D31");

            Merge(ws, "A1:G1", "Top 10 Posiciones por Valuación",
                13, bold: true, fg: Hex("1F4E79"), bg: Hex("DEEAF1"), halign: ExcelHorizontalAlignment.Center);
            ws.Row(1).Height = 26;

            var hdrs = new[] { "#","Símbolo","Descripción","Tipo","Valorizado $","Ganancia $","Ganancia %" };
            for (int c = 0; c < hdrs.Length; c++)
                SetHeader(ws.Cells[2, c + 1], hdrs[c]);
            ws.Row(2).Height = 22;

            var top = items.OrderByDescending(p => p.Valorizado).Take(10).ToList();

            for (int i = 0; i < top.Count; i++)
            {
                int row = i + 3;
                var p = top[i];
                var tipo = NormTipo(p.Tipo);
                var (bg, fg, _) = Colores(tipo);

                // Medalla top 3
                Color rankBg = i == 0 ? Hex("FFD700") : i == 1 ? Hex("C0C0C0") : i == 2 ? Hex("CD7F32") : bg;
                Color rankFg = i < 3 ? Hex("1A1A1A") : fg;

                var rank = ws.Cells[row, 1];
                rank.Value = i + 1;
                rank.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rank.Style.Fill.BackgroundColor.SetColor(rankBg);
                rank.Style.Font.Bold = true;
                rank.Style.Font.Color.SetColor(rankFg);
                rank.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells[row, 2].Value = p.Simbolo;
                ws.Cells[row, 3].Value = p.Descripcion;
                ws.Cells[row, 4].Value = tipo;
                ws.Cells[row, 5].Value = p.Valorizado;
                ws.Cells[row, 6].Value = p.GananciaDinero;
                ws.Cells[row, 7].Value = p.GananciaPorcentaje / 100m;

                var rng = ws.Cells[row, 2, row, 7];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(bg);
                rng.Style.Font.Color.SetColor(fg);
                rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                rng.Style.Border.Bottom.Color.SetColor(Hex("D9D9D9"));
                ws.Row(row).Height = 18;

                ws.Cells[row, 5].Style.Numberformat.Format = "#,##0.00";
                ws.Cells[row, 6].Style.Numberformat.Format = "#,##0.00";
                ws.Cells[row, 7].Style.Numberformat.Format = "0.00%";

                SemaforoBold(ws.Cells[row, 6], p.GananciaDinero);
                Semaforo(ws.Cells[row, 7], p.GananciaPorcentaje);
            }

            ws.Column(1).Width = 5;  ws.Column(2).Width = 12; ws.Column(3).Width = 36;
            ws.Column(4).Width = 18; ws.Column(5).Width = 17; ws.Column(6).Width = 17;
            ws.Column(7).Width = 12;

            // Gráfico de barras horizontales — Top 10 valuación
            int de = top.Count + 2;
            var bar = (ExcelBarChart)ws.Drawings.AddChart("Top10Bar", eChartType.BarClustered);
            bar.Series.Add(ws.Cells[3, 5, de, 5], ws.Cells[3, 2, de, 2]);
            bar.Series[0].Header = "Valorizado $";
            bar.Title.Text = "Top 10 por Valuación";
            bar.Title.Font.Size = 11;
            bar.Legend.Remove();
            bar.SetPosition(2, 0, 8, 0);
            bar.SetSize(400, 310);
        }

        // ═══════════════════════════════════════════════════
        //  HELPERS
        // ═══════════════════════════════════════════════════

        private static string NormTipo(string? tipo)
        {
            if (string.IsNullOrWhiteSpace(tipo)) return "Otro";
            if (tipo.Contains("CEDEAR",    StringComparison.OrdinalIgnoreCase)) return "CEDEAR";
            if (tipo.Contains("ACCION",    StringComparison.OrdinalIgnoreCase)) return "Acciones";
            if (tipo.Contains("TITULO",    StringComparison.OrdinalIgnoreCase) ||
                tipo.Contains("BONO",      StringComparison.OrdinalIgnoreCase)) return "Títulos Públicos";
            if (tipo.Contains("LETRA",     StringComparison.OrdinalIgnoreCase)) return "Letras";
            if (tipo.Contains("FCI",       StringComparison.OrdinalIgnoreCase) ||
                tipo.Contains("FONDO",     StringComparison.OrdinalIgnoreCase)) return "FCI";
            if (tipo.Contains("OBLIGACION",StringComparison.OrdinalIgnoreCase) ||
                tipo.Equals("ON",          StringComparison.OrdinalIgnoreCase)) return "ON";
            return tipo;
        }

        private static void SetHeader(ExcelRange cell, string text)
        {
            cell.Value = text;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(Hex("1F4E79"));
            cell.Style.Font.Color.SetColor(Color.White);
            cell.Style.Font.Bold = true;
            cell.Style.Font.Size = 10;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment   = ExcelVerticalAlignment.Center;
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Hex("16375E"));
        }

        private static void Merge(ExcelWorksheet ws, string addr, string text, int fontSize,
            bool bold = false, Color? fg = null, Color? bg = null,
            ExcelHorizontalAlignment halign = ExcelHorizontalAlignment.Left)
        {
            var cell = ws.Cells[addr];
            cell.Merge = true;
            cell.Value = text;
            cell.Style.Font.Size = fontSize;
            cell.Style.Font.Bold = bold;
            if (fg.HasValue) cell.Style.Font.Color.SetColor(fg.Value);
            if (bg.HasValue && bg.Value != Color.Transparent)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(bg.Value);
            }
            cell.Style.HorizontalAlignment = halign;
            cell.Style.VerticalAlignment   = ExcelVerticalAlignment.Center;
        }

        private static void Semaforo(ExcelRange cell, decimal valor)
            => cell.Style.Font.Color.SetColor(valor >= 0 ? Hex("375623") : Hex("9C0006"));

        private static void SemaforoBold(ExcelRange cell, decimal valor)
        {
            cell.Style.Font.Bold = true;
            Semaforo(cell, valor);
        }

        private static Color Hex(string h)
        {
            h = h.TrimStart('#');
            return Color.FromArgb(
                Convert.ToInt32(h[0..2], 16),
                Convert.ToInt32(h[2..4], 16),
                Convert.ToInt32(h[4..6], 16));
        }

        private static Color Blend(Color a, Color b, float t) => Color.FromArgb(
            (int)(a.R + (b.R - a.R) * t),
            (int)(a.G + (b.G - a.G) * t),
            (int)(a.B + (b.B - a.B) * t));
    }
}
