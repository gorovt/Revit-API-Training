using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddinHolaMundo
{
    public class ParamGrid
    {
        public int id { get; set; }
        public string nombre { get; set; }
        public string valor { get; set; }

        public ParamGrid()
        {
            id = 0;
            nombre = string.Empty;
            valor = string.Empty;
        }

        public ParamGrid(int id, string nombre, string valor)
        {
            this.id = id;
            this.nombre = nombre;
            this.valor = valor;
        }
    }
}
