# EPPLUS-Importar
importar y exportar desde una clase generica(DTO) con epplus 

## 📚 Librerias
 - Epplus 7.1

## Importar datos desde un archivo excel a una clase 
```c#
List<Formato> datos = new EPPlusExcel<Formato>("file", 2)
.MapearColumna(x => x.Id, "A", 1)
.MapearColumna(x => x.Nombre, "B", 1)
.MapearColumna(x => x.Edad, "C", 1)
.ExtraerDatos();
Console.WriteLine("Hello, World!"); 
```
## Exportar a bytes
```c#
List<Formato> lista = new List<Formato>();
lista.Add(new Formato { Id = 1, Nombre = "ivan", Edad = 19 });
lista.Add(new Formato { Id = 2, Nombre = "jorge", Edad = 20});

var exportar =await new EPPlusExcel<Formato>(lista,"file",2)
.MapearColumna(x => x.Id, "A")
.MapearColumna(x => x.Nombre, "B")
.MapearColumna(x => x.Edad, "C")
.ExportarDatos();
```
