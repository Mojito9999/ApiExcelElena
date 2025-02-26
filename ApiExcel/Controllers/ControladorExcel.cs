using Microsoft.AspNetCore.Mvc; // Importa el namespace necesario para crear controladores de API
using OfficeOpenXml; // Importa la biblioteca EPPlus para manipular archivos Excel
using System.Collections.Generic; // Importa para usar listas genéricas
using System.IO; // Importa para manipular archivos y directorios
using System.Linq; // Importa para usar LINQ para consultas
using System.Threading.Tasks; // Importa para manejar tareas asíncronas




namespace ApiExcel.Controllers
{
    [ApiController]
    [Route("api/[controller]")]

    // Controlador excel hereda de ControllerBase
    public class ControladorExcel : ControllerBase
    {
        // Ponemos la ruta del archivo excel
        private const string ExcelFilePath = @"C:\DatosExcel\archivo.xlsx";

        // METODO SOLICITUD HTTP GET
        [HttpGet]
        //SOLICITUD PARA VER LOS DATOS DEL EXCEL
        public IActionResult GetExcelData()
        {
            // COMPROBACION DE SI EXISTE EL ARCHIVO
            if (!System.IO.File.Exists(ExcelFilePath))
            {
                return NotFound("El archivo no existe.");
            }

            try
            {
                // Guardamos los empleados en una lista
                var empleados = new List<Empleado>();

                // Usamos EPPlus para abrir el excel
                using (var package = new ExcelPackage(new FileInfo(ExcelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // Obtiene hoja del libro excel
                    var rowCount = worksheet.Dimension.Rows; // Obtienemos el número de filas



                    // ITERAMOS LAS FILAS
                    for (int row = 1; row <= rowCount; row++)
                    {
                        // Creamos nuevo empleado
                        var empleado = new Empleado
                        {
                            // Leemos el nombre del empleado en la columna 2
                            Nombre = worksheet.Cells[row, 2].Text,
                            Salario = 0
                        };

                        // Leemos el salario del empleado en la columna 4
                        var salarios = worksheet.Cells[row, 4].Text;

                        // Verifica si la celda no está vacía y si se puede convertir a decimal
                        if (!string.IsNullOrWhiteSpace(salarios) && decimal.TryParse(salarios, out decimal salario))
                        {
                            //Asigno el salario al obj
                            empleado.Salario = salario;
                            //Si el salario es mayor a .....
                            empleado.CobranMas7500 = empleado.Salario > 7500 ? "Sí" : "No";

                            //Añadimos el empleado a la lista
                            empleados.Add(empleado);


                            // Los guardamos con el package de abajo del for pero aqi guardamos si o no
                            //dependiendo de si son más de 7500 con el filtro de arriba como en la api
                            worksheet.Cells[row, 7].Value = empleado.CobranMas7500;
                          

                        }
                        else
                        {
                            // Registra un error si no vale el salario
                            Log.Instance.LogError($"El salario en la fila {row} no es un número válido: '{salarios}'");
                        }
                    }

                    //GUARDO LOS CAMBIOS EN EL EXCEL
                    package.Save();

                }

                // Ordenamos la lista descendientemente por nombre
                var listaEmpleados = empleados.OrderByDescending(e => e.Nombre).ToList();
                //Devuelve la lista en json
                return Ok(listaEmpleados);
            }
            catch (Exception ex)
            {
                //Excepciones para la carpeta log
                //LogError(ex);
                //Nueva clase usando singleton para registrar las excepciones en vez de en la carpeta con el metodo de abajo
                Log.Instance.LogError(ex.Message);
                //Devuelve si hay problemas
                return StatusCode(500, "Error al procesar el archivo.");
            }
        }

        // METODO ERRORES DEL LOG PARA VERLO EN LA CARPETA
        /*private void LogError(Exception ex)
        {
            // Define la ruta del archivo de log con la fecha actual
            var logPath = Path.Combine(@"C:\Logs", $"{DateTime.Now:yyyy-MM-dd}.log");
            // Abre el archivo de log para agregar nuevas entradas
            using (var writer = new StreamWriter(logPath, true))
            {
                writer.WriteLine($"{DateTime.Now}: {ex.Message}"); // Escribe la fecha y el mensaje de error
                writer.WriteLine(ex.StackTrace); // Escribe la pila de llamadas para depuración
            }
        }*/
    }

    // CLASE EMPLEADO
    public class Empleado
    {
        public string Nombre { get; set; }
        public decimal Salario { get; set; }
        public string CobranMas7500 { get; set; }
    }
}
