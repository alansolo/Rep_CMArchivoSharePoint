using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CMArchivoSharePoint.Model
{
    public class Archivo
    {
        public string Area { get; set; }
        public string Departamento { get; set; }
        public string TipoDocumento { get; set; }
        public string DepartamentoCodigo { get; set; }
        public string Codigo { get; set; }
        public string NombreDocumento { get; set; }
        public string DescripcionDocumento { get; set; }
        public string NumeroRevision { get; set; }
        public string FCambioFijo { get; set; }
        public string FCambioFrecuente { get; set; }
        public string SBU { get; set; }
        public string Cliente { get; set; }
    }
}
