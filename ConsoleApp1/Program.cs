// See https://aka.ms/new-console-template for more information

using ConsoleApp1;

/*List<Formato> datos = new EPPlusExcel<Formato>("file", 2)
.MapearColumna(x => x.Id, "A", 1)
.MapearColumna(x => x.Nombre, "B", 1)
.MapearColumna(x => x.Edad, "C", 1)
.ExtraerDatos();
Console.WriteLine("Hello, World!"); */


List<Formato> lista = new List<Formato>();
lista.Add(new Formato { Id = 1, Nombre = "ivan", Edad = 19 });
lista.Add(new Formato { Id = 2, Nombre = "jorge", Edad = 20});

var exportar =await new EPPlusExcel<Formato>(lista,"file",2)
.MapearColumna(x => x.Id, "A", 1)
.MapearColumna(x => x.Nombre, "B", 1)
.MapearColumna(x => x.Edad, "C", 1)
.ExportarDatos();
File.WriteAllBytes("C:\\Users\\USUARIO\\Desktop\\mapeoCelda\\miarchivo.xlsx", exportar);

