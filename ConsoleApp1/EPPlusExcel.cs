using Microsoft.Win32;
using OfficeOpenXml;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;
using System.Web.Mvc;

namespace ConsoleApp1
{
    public class EPPlusExcel<TModel>
    {
        public EPPlusExcel(string file, int filaInicio)
        {
            _filaInicio = filaInicio;
            fileHttpPosted = file;
        }
        public EPPlusExcel(List<TModel> lista,string file, int filaInicio)
        {
            _filaInicio = filaInicio;
            fileHttpPosted = file;
            RegistrosExcel = lista;
        }
        public int _filaInicio = 0;
        public string fileHttpPosted;
        public List<MapeoColumnaExcel> Columnas { get; set; } = new List<MapeoColumnaExcel>();
        public List<TModel> RegistrosExcel { get; set; } = new List<TModel>();

    }

    public static class ImportarExcelExtensions
    {
        public static EPPlusExcel<TModel> MapearColumna<TModel, Tproperty>(this EPPlusExcel<TModel> modelo, Expression<Func<TModel,
            Tproperty>> propiedad, string columna, int fila) where TModel : new()
        {
            string nombrePropiedad = ExpressionHelper.GetExpressionText(propiedad);
            Type tipo = propiedad.Body.Type;
            modelo.Columnas.Add(new MapeoColumnaExcel()
            {
                Propiedad = nombrePropiedad,
                Fila = fila,
                Columna = columna,
                Tipo = tipo,
            });
            return modelo;
        }

        public static List<TModel> ExtraerDatos<TModel>(this EPPlusExcel<TModel> importar) where TModel : new()
        {
            if (importar.fileHttpPosted is null)
            {
                throw new ArgumentNullException("no se especifico ningun archivo para importar");
            }
            if (importar.fileHttpPosted is null)
            {
                throw new ArgumentNullException("solo se permiten archivos de tipo excel para importar");
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string directorioBase = AppDomain.CurrentDomain.BaseDirectory;
            // Concatena la ruta del directorio base con el nombre del archivo
            string rutaArchivo = Path.Combine(directorioBase, "prueba.xlsx");
            using (var package = new ExcelPackage("C:\\Users\\USUARIO\\Desktop\\mapeoCelda\\ConsoleApp1\\prueba.xlsx"))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets.First();
                //int rowCount = worksheet.Dimension.End.Column;
                int totalFilas = worksheet.Dimension.Rows;
                for (int fila = importar._filaInicio; fila <= totalFilas; fila++)
                {
                    var data = worksheet.ExtraerDatosCampo(new TModel(), importar.Columnas, fila);
                    importar.RegistrosExcel.Add(data);
                }
            }

            return importar.RegistrosExcel;
        }

        private static TModel ExtraerDatosCampo<TModel>(this ExcelWorksheet worksheet, TModel modelo, List<MapeoColumnaExcel> Columnas, int fila)
            where TModel : new()
        {
            var tprops = (new TModel())
              .GetType()
              .GetProperties()
              .ToList();
            var newModel = new TModel();
            foreach (MapeoColumnaExcel columna in Columnas) {
                var prop = tprops.First(p => p.Name == columna.Propiedad);
                if(columna.Tipo == typeof(int))
                {
                    int valueFila = 0;
                    string valor = worksheet.Cells[$"{columna.Columna}{fila}"].Value.ToString() ?? "";
                    int.TryParse(valor, out valueFila);
                    prop.SetValue(newModel, valueFila);
                }
                else if (columna.Tipo == typeof(string))
                {
                    string valueFila = worksheet.Cells[$"{columna.Columna}{fila}"].Value.ToString() ?? "";
                    prop.SetValue(newModel, valueFila);
                }    
            }

            return newModel;
        }

    }


}
