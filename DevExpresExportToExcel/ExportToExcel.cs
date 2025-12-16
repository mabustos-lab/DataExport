using ClosedXML.Excel;
using System.Data;
using System.Diagnostics;

namespace DataExport
{
    public static class ExportToExcel
    {
        /// <summary>
        /// Exporta a Excel un <see cref="DataTable"/>, usando <see cref="SaveFileDialog"/>
        /// </summary>
        /// <param name="dataSource">DataTable que se requiere exportar</param>
        /// <param name="filePath">Ruta del archivo excel.</param>
        /// <param name="openFile">Indica se se abre el archivo al finalizar la exporación</param>
        /// <returns>India</returns>
        public static bool ToExcel(this DataTable dataSource, string filePath, bool openFile = true)
        {
            // Validaciones básicas
            if (dataSource == null)
                throw new ArgumentNullException(nameof(dataSource), "El DataTable no puede ser nulo");
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("La ruta del archivo no puede estar vacía", nameof(filePath));
            bool result = false;
            try
            {
                Export(dataSource, filePath);
                if (openFile)
                    OpenFile(filePath);
                result = true;
            }
            catch (Exception ex )
            {
                Debug.Write($"Ocurrio un problema al exportar los datos: {ex.Message}");
            }
            return result;
        }
        /// <summary>
        /// Abre el archivo con la aplicación predeterminada del sistema.
        /// </summary>
        private static void OpenFile(string filePath)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = filePath;
                process.StartInfo.UseShellExecute = true;
                process.Start();
            }
            catch (Exception ex)
            {
                Debug.Write($"Ocurrio un problema al abrir el archivo: {ex.Message}");
                throw;
            }
        }
        /// <summary>
        /// Exporta un DataTable a un archivo Excel usando ClosedXML
        /// </summary>
        /// <param name="dt">DataTable con los datos a exportar</param>
        /// <param name="filePath">Ruta completa del archivo Excel a crear</param>
        private static void Export(DataTable dt, string filePath)
        {
            // Crear un nuevo libro de trabajo
            using (var workbook = new XLWorkbook())
            {
                // Usar TableName si existe, de lo contrario usar "Hoja1"
                string sheetName = !string.IsNullOrWhiteSpace(dt.TableName)
                    ? GetValidSheetName(dt.TableName)
                    : "Hoja1";
                // Crear la hoja de trabajo
                var worksheet = workbook.Worksheets.Add(sheetName);

                // Escribir encabezados
                int row = 1;
                int col = 1;

                foreach (DataColumn column in dt.Columns)
                {
                    // Usar Caption si está configurado, de lo contrario usar el nombre de la columna
                    string headerText = !string.IsNullOrWhiteSpace(column.Caption)
                        ? column.Caption
                        : column.ColumnName;
                    // Escribir el encabezado
                    worksheet.Cell(row, col).Value = headerText;

                    // Aplicar formato al encabezado
                    worksheet.Cell(row, col).Style.Font.Bold = true;
                    worksheet.Cell(row, col).Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell(row, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    // Obtener la descripción de la columna para usar como tooltip
                    string description = column.GetDescription();

                    if (!string.IsNullOrWhiteSpace(description))
                    {
                        // Agregar tooltip como comentario (aparece al pasar el mouse)
                        var validation= worksheet.Cell(row, col).GetDataValidation();
                        validation.InputTitle = $"Información: {headerText}";
                        validation.InputMessage= description;
                        // Permitir cualquier valor (solo queremos el mensaje)
                        validation.IgnoreBlanks = true;
                    }
                    col++;
                }
                // Escribir datos
                row = 2; // Empezar en la fila 2 (debajo de los encabezados)
                foreach (DataRow dataRow in dt.Rows)
                {
                    col = 1;
                    foreach (DataColumn column in dt.Columns)
                    {
                        var cell = worksheet.Cell(row, col);

                        // Manejar valores nulos o DBNull
                        if (dataRow[column] == DBNull.Value || dataRow[column] == null)
                        {
                            cell.Value = string.Empty;
                        }
                        else
                        {
                            // Usar el método corregido para convertir valores
                            SetCellValue(cell, dataRow[column], column.DataType);
                        }
                        // Aplicar borde a las celdas de datos
                        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        col++;
                    }
                    row++;

                }

                // Ajustar automáticamente el ancho de las columnas
                worksheet.Columns().AdjustToContents();

                // Guardar el archivo
                workbook.SaveAs(filePath);
            }
        }

        /// <summary>
        /// Establece el valor de una celda manejando diferentes tipos de datos
        /// </summary>
        private static void SetCellValue(IXLCell cell, object value, Type dataType)
        {
            if (value == null)
            {
                cell.Value = string.Empty;
                return;
            }

            // Manejar tipos específicos de manera explícita
            if (dataType == typeof(DateTime) || dataType == typeof(DateTime?))
            {
                if (value is DateTime dateTimeValue)
                {
                    cell.Value = dateTimeValue;
                    cell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
                }
                else if (DateTime.TryParse(value.ToString(), out DateTime parsedDate))
                {
                    cell.Value = parsedDate;
                    cell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
                }
                else
                {
                    cell.Value = value.ToString();
                }
            }
            else if (dataType == typeof(bool) || dataType == typeof(bool?))
            {
                if (value is bool boolValue)
                {
                    cell.Value = boolValue;
                }
                else if (bool.TryParse(value.ToString(), out bool parsedBool))
                {
                    cell.Value = parsedBool;
                }
                else
                {
                    cell.Value = value.ToString();
                }
            }
            else if (dataType == typeof(int) || dataType == typeof(int?))
            {
                if (value is int intValue)
                {
                    cell.Value = intValue;
                }
                else if (int.TryParse(value.ToString(), out int parsedInt))
                {
                    cell.Value = parsedInt;
                }
                else
                {
                    cell.Value = value.ToString();
                }
            }
            else if (dataType == typeof(long) || dataType == typeof(long?))
            {
                if (value is long longValue)
                {
                    cell.Value = longValue;
                }
                else if (long.TryParse(value.ToString(), out long parsedLong))
                {
                    cell.Value = parsedLong;
                }
                else
                {
                    cell.Value = value.ToString();
                }
            }
            else if (dataType == typeof(decimal) || dataType == typeof(decimal?))
            {
                if (value is decimal decimalValue)
                {
                    cell.Value = decimalValue;
                }
                else if (decimal.TryParse(value.ToString(), out decimal parsedDecimal))
                {
                    cell.Value = parsedDecimal;
                }
                else
                {
                    cell.Value = value.ToString();
                }
            }
            else if (dataType == typeof(double) || dataType == typeof(double?))
            {
                if (value is double doubleValue)
                {
                    cell.Value = doubleValue;
                }
                else if (double.TryParse(value.ToString(), out double parsedDouble))
                {
                    cell.Value = parsedDouble;
                }
                else
                {
                    cell.Value = value.ToString();
                }
            }
            else if (dataType == typeof(float) || dataType == typeof(float?))
            {
                if (value is float floatValue)
                {
                    cell.Value = (double)floatValue; // ClosedXML usa double para números
                }
                else if (float.TryParse(value.ToString(), out float parsedFloat))
                {
                    cell.Value = (double)parsedFloat;
                }
                else
                {
                    cell.Value = value.ToString();
                }
            }
            else if (dataType == typeof(short) || dataType == typeof(short?))
            {
                if (value is short shortValue)
                {
                    cell.Value = shortValue;
                }
                else if (short.TryParse(value.ToString(), out short parsedShort))
                {
                    cell.Value = parsedShort;
                }
                else
                {
                    cell.Value = value.ToString();
                }
            }
            else if (dataType == typeof(byte) || dataType == typeof(byte?))
            {
                if (value is byte byteValue)
                {
                    cell.Value = byteValue;
                }
                else if (byte.TryParse(value.ToString(), out byte parsedByte))
                {
                    cell.Value = parsedByte;
                }
                else
                {
                    cell.Value = value.ToString();
                }
            }
            else if (dataType == typeof(Guid) || dataType == typeof(Guid?))
            {
                if (value is Guid guidValue)
                {
                    cell.Value = guidValue.ToString();
                }
                else if (Guid.TryParse(value.ToString(), out Guid parsedGuid))
                {
                    cell.Value = parsedGuid.ToString();
                }
                else
                {
                    cell.Value = value.ToString();
                }
            }
            else
            {
                // Para cualquier otro tipo, usar como string
                cell.Value = value.ToString();
            }
        }

        /// <summary>
        /// Obtiene un nombre válido para hoja de Excel
        /// </summary>
        private static string GetValidSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Hoja1";

            // Excel tiene restricciones para nombres de hojas:
            // - Máximo 31 caracteres
            // - No puede contener: : \ / ? * [ ]
            // - No puede empezar o terminar con apóstrofe
            // - No puede estar vacío

            string validName = name;

            // Limitar a 31 caracteres
            if (validName.Length > 31)
                validName = validName.Substring(0, 31);

            // Reemplazar caracteres no permitidos
            char[] invalidChars = new char[] { ':', '\\', '/', '?', '*', '[', ']' };
            foreach (char c in invalidChars)
            {
                validName = validName.Replace(c, '_');
            }

            // Eliminar apóstrofes al inicio y final
            validName = validName.Trim('\'');

            // Si queda vacío, usar nombre por defecto
            if (string.IsNullOrWhiteSpace(validName))
                validName = "Hoja1";

            return validName;
        }
    }
}
