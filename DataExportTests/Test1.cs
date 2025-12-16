using System.Data;
using DataExport;

namespace DataExportTests
{
    [TestClass]
    public sealed class Test1
    {
        [TestMethod]
        public void TestExportSimpleTestDataExcel()
        {
            DataTable productos = CreateSimpleTestDataTable();
            productos.ToExcel(@"C:\Temp\Productos_Test.xlsx",true);
        }
        [TestMethod]
        public void TestExportTestDataExcel()
        {
            DataTable productos = CreateTestDataTable();
            productos.ToExcel(@"C:\Temp\Productos_Test1.xlsx", true);
        }

        /// <summary>
        /// Crea un DataTable de ejemplo para pruebas con diferentes tipos de datos
        /// </summary>
        public static DataTable CreateTestDataTable()
        {
            // Crear DataTable con nombre personalizado
            DataTable dt = new DataTable("Empleados_Prueba");

            // Agregar columnas con diferentes tipos de datos y descripciones
            dt.AddColumn<int>("ID",
                "ID Empleado",
                "Identificador único del empleado",
                excludeColumnFlag: true);

            dt.AddColumn<string>("Nombre",
                "Nombre Completo",
                "Ingrese el nombre completo del empleado");

            dt.AddColumn<string>("Apellido",
                "Apellidos",
                "Apellidos del empleado");

            dt.AddColumn<DateTime>("FechaNacimiento",
                "Fecha de Nacimiento",
                "Fecha de nacimiento (formato: DD/MM/AAAA)");

            dt.AddColumn<DateTime>("FechaIngreso",
                "Fecha de Ingreso",
                "Fecha en que el empleado ingresó a la empresa");

            dt.AddColumn<decimal>("Salario",
                "Salario Base",
                "Salario mensual bruto en moneda local");

            dt.AddColumn<double>("Comision",
                "Porcentaje Comisión",
                "Porcentaje de comisión sobre ventas (ej: 0.15 = 15%)");

            dt.AddColumn<bool>("Activo",
                "Empleado Activo",
                "Indica si el empleado está activo en la empresa");

            dt.AddColumn<int?>("DepartamentoID",
                "ID Departamento",
                "Identificador del departamento asignado");

            dt.AddColumn<string>("Departamento",
                "Nombre Departamento",
                "Nombre del departamento")
                .AllowDBNull = true;

            dt.AddColumn<string>("Email",
                "",
                "Dirección de correo electrónico corporativo")
                .AllowDBNull=true;

            dt.AddColumn<string>("Telefono",
                "Teléfono Contacto",
                "Número de teléfono de contacto")
                .AllowDBNull = true;

            dt.AddColumn<DateTime?>("FechaBaja",
                "Fecha de Baja",
                "Fecha de baja (solo si el empleado ya no está activo)");

            dt.AddColumn<decimal?>("Bonificacion",
                "Bonificación Anual",
                "Bonificación anual si aplica (opcional)");

            dt.AddColumn<Guid>("UUID",
                "Identificador Único",
                "Identificador único global (GUID) del registro");

            // Agregar filas de datos de ejemplo
            dt.Rows.Add(
                1,
                "Juan",
                "Pérez García",
                new DateTime(1985, 3, 15),
                new DateTime(2010, 5, 20),
                2500.50m,
                0.10,
                true,
                101,
                "Ventas",
                "juan.perez@empresa.com",
                "+34 600 123 456",
                DBNull.Value,
                1500.75m,
                Guid.NewGuid()
            );

            dt.Rows.Add(
                2,
                "María",
                "López Fernández",
                new DateTime(1990, 7, 22),
                new DateTime(2015, 9, 10),
                3200.00m,
                0.15,
                true,
                102,
                "Marketing",
                "maria.lopez@empresa.com",
                "+34 600 234 567",
                DBNull.Value,
                2000.00m,
                Guid.NewGuid()
            );

            dt.Rows.Add(
                3,
                "Carlos",
                "Martínez Ruiz",
                new DateTime(1978, 11, 5),
                new DateTime(2005, 2, 28),
                4200.75m,
                0.12,
                true,
                103,
                "IT",
                "carlos.martinez@empresa.com",
                "+34 600 345 678",
                DBNull.Value,
                3000.50m,
                Guid.NewGuid()
            );

            dt.Rows.Add(
                4,
                "Ana",
                "Gómez Sánchez",
                new DateTime(1995, 1, 30),
                new DateTime(2020, 3, 15),
                1800.25m,
                0.08,
                false,
                101,
                "Ventas",
                "ana.gomez@empresa.com",
                "+34 600 456 789",
                new DateTime(2023, 6, 30),
                DBNull.Value,
                Guid.NewGuid()
            );

            dt.Rows.Add(
                5,
                "Pedro",
                "Rodríguez Martín",
                new DateTime(1988, 9, 12),
                new DateTime(2018, 11, 5),
                2900.00m,
                0.11,
                true,
                104,
                "Recursos Humanos",
                "pedro.rodriguez@empresa.com",
                "+34 600 567 890",
                DBNull.Value,
                1800.00m,
                Guid.NewGuid()
            );

            // Agregar fila con valores nulos para pruebas
            dt.Rows.Add(
                6,
                "Laura",
                "Hernández",
                new DateTime(1992, 4, 18),
                new DateTime(2022, 1, 10),
                2100.00m,
                0.09,
                true,
                DBNull.Value,
                DBNull.Value,
                DBNull.Value,
                DBNull.Value,
                DBNull.Value,
                DBNull.Value,
                Guid.NewGuid()
            );

            return dt;
        }

        /// <summary>
        /// Crea un DataTable de ejemplo más simple para pruebas básicas
        /// </summary>
        public static DataTable CreateSimpleTestDataTable()
        {
            DataTable dt = new DataTable("Productos");

            dt.AddColumn<int>("Codigo", "Código Producto", "Código único del producto");
            dt.AddColumn<string>("Nombre", "Nombre Producto", "Nombre descriptivo del producto");
            dt.AddColumn<string>("Categoria", "Categoría", "Categoría a la que pertenece el producto");
            dt.AddColumn<decimal>("Precio", "Precio Unitario", "Precio de venta unitario (sin IVA)");
            dt.AddColumn<int>("Stock", "Existencias", "Cantidad disponible en inventario");
            dt.AddColumn<DateTime>("UltimaReposicion", "Última Reposición", "Fecha de la última reposición de inventario");
            dt.AddColumn<bool>("Disponible", "Disponible Venta", "Indica si el producto está disponible para venta");

            dt.Rows.Add(1001, "Laptop Gaming", "Electrónica", 1299.99m, 15, new DateTime(2024, 1, 15), true);
            dt.Rows.Add(1002, "Mouse Inalámbrico", "Electrónica", 49.99m, 120, new DateTime(2024, 2, 10), true);
            dt.Rows.Add(1003, "Silla Ergonómica", "Oficina", 349.50m, 8, new DateTime(2024, 1, 30), true);
            dt.Rows.Add(1004, "Monitor 27\" 4K", "Electrónica", 599.99m, 0, new DateTime(2023, 12, 5), false);
            dt.Rows.Add(1005, "Teclado Mecánico", "Electrónica", 89.99m, 45, new DateTime(2024, 2, 20), true);

            return dt;
        }
    }
}
