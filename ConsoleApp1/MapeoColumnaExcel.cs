using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace ConsoleApp1
{
    public class MapeoColumnaExcel
    {
        public string Propiedad { get; set; }
        public Type Tipo { get; set; }
        public string Columna { get; set; }
        public int Fila { get; set; }

    }
}
