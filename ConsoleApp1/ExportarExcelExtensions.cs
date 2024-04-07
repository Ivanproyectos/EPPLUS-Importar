using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Reflection;
using Microsoft.Win32;

namespace ConsoleApp1
{
    public static class ExportarExcelExtensions
    {
        public static async Task<byte[]> ExportarDatos<TModel>(this EPPlusExcel<TModel> modeloDinamico) where TModel : new()
        {
            if (modeloDinamico.fileHttpPosted is null)
            {
                throw new ArgumentNullException("no se especifico ningun archivo para importar");
            }
            if (modeloDinamico.fileHttpPosted is null)
            {
                throw new ArgumentNullException("solo se permiten archivos de tipo excel para importar");
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var memory = new MemoryStream())
            {
                using (var package = new ExcelPackage("C:\\Users\\USUARIO\\Desktop\\mapeoCelda\\ConsoleApp1\\plantilla.xlsx"))
                {
                    var workbook = package.Workbook;
                    var worksheet = workbook.Worksheets.First();
                    int totalFilas = worksheet.Dimension.Rows;
                    //for (int fila = modeloDinamico._filaInicio; fila <= totalFilas; fila++)
                    //{
                    //    var data = worksheet.ExportarDatosEnCampo(new TModel(), modeloDinamico.Columnas, fila);
                    //    modeloDinamico.RegistrosExcel.Add(data);
                    //}
                    int filaInicio = modeloDinamico._filaInicio;
                    int contador = 0;
                    modeloDinamico.RegistrosExcel.ForEach(registro =>
                    {
                        worksheet.ExportarDatosEnCampo(registro, modeloDinamico.Columnas, filaInicio);
                        filaInicio++;
                    });
                    await package.SaveAsAsync(memory);
                    return memory.ToArray(); 
                }
            }
            //return importar.RegistrosExcel;
        }
        private static void ExportarDatosEnCampo<TModel>(this ExcelWorksheet worksheet, TModel registro, List<MapeoColumnaExcel> mapeoColumnas, int fila)
          where TModel : new()
        {

            foreach (MapeoColumnaExcel columna in mapeoColumnas)
            {
                PropertyInfo[] props = registro?.GetType().GetProperties() ?? new PropertyInfo[0];
                PropertyInfo propertyInfo = props.First(p => p.Name == columna.Propiedad);
                object propertyValue = propertyInfo?.GetValue(registro) ?? new object();
                worksheet.Cells[$"${columna.Columna}{fila}"].Value = propertyValue;
            }
     
        }

    }
}
