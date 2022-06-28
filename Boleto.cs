using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Boletos
{
    public class Boleto
    {
        public string CodigoDeBarras { get; set; }
        public string Valor { get; set; }
        public string Cnpj { get; set; }
        public string Eolica { get; set; }
    }
}
