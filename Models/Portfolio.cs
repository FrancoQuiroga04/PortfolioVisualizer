using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IOLPortfolio.Models
{
    public class Portfolio
    {
        public string Simbolo { get; set; }
        public string Descripcion { get; set; }
        public string Tipo { get; set; }
        public string Mercado { get; set; }
        public string Moneda { get; set; }
        public decimal Cantidad { get; set; }
        public decimal UltimoPrecio { get; set; }
        public decimal Ppc { get; set; }
        public decimal Valorizado { get; set; }
        public decimal GananciaDinero { get; set; }
        public decimal GananciaPorcentaje { get; set; }
        public decimal VariacionDiaria { get; set; }
    }
}
