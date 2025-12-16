using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataExport
{
    /// <summary>
    /// Proporciona métodos de extensión para la clase <see cref="DataTable"/>.
    /// </summary>
    public static class DataTableExt
    {
        /// <summary>
        /// Clave usada para marcar columnas que deben excluirse en ciertos procesos.
        /// </summary>
        private const string ExcludeColumnKey = "ExcludeColumn";

        /// <summary>
        /// Clave usada para almacenar la descripción de una columna.
        /// </summary>
        private const string ColumnDescriptionKey = "ColumnDescription";

        /// <summary>
        /// Agrega una nueva columna al <see cref="DataTable"/> con el tipo especificado.
        /// </summary>
        /// <typeparam name="TType">Tipo de datos de la columna.</typeparam>
        /// <param name="dt"><see cref="DataTable"/> al que se agregará la columna.</param>
        /// <param name="fieldName">Nombre del campo de la columna.</param>
        /// <param name="captionColumn">Texto a mostrar como cabecera (opcional).</param>
        /// <param name="description">Descripción de la columna (opcional).</param>
        /// <param name="excludeColumnFlag">Indica si la columna debe excluirse en ciertos procesos (opcional).</param>
        /// <returns>La columna recién creada o <c>null</c> si <paramref name="dt"/> es <c>null</c>.</returns>
        public static DataColumn AddColumn<TType>(
            this DataTable dt,
            string fieldName,
            string captionColumn = "",
            string description = "",
            bool excludeColumnFlag = false)
        {
            return dt.AddColumn(typeof(TType), fieldName, captionColumn, description, excludeColumnFlag);
        }

        /// <summary>
        /// Agrega una nueva columna al <see cref="DataTable"/> con el tipo especificado.
        /// </summary>
        /// <param name="dt"><see cref="DataTable"/> al que se agregará la columna.</param>
        /// <param name="dataType">Tipo de datos de la columna.</param>
        /// <param name="fieldName">Nombre del campo de la columna.</param>
        /// <param name="captionColumn">Texto a mostrar como cabecera (opcional).</param>
        /// <param name="description">Descripción de la columna (opcional).</param>
        /// <param name="excludeColumnFlag">Indica si la columna debe excluirse en ciertos procesos (opcional).</param>
        /// <returns>La columna recién creada o <c>null</c> si <paramref name="dt"/> es <c>null</c>.</returns>
        public static DataColumn AddColumn(
            this DataTable dt,
            Type dataType,
            string fieldName,
            string captionColumn = "",
            string description = "",
            bool excludeColumnFlag = false)
        {
            if (dt == null)
                throw new ArgumentNullException(nameof(dt));

            bool allowDBNull = false;
            Type underlyingType = Nullable.GetUnderlyingType(dataType);

            // Si es un tipo nullable, se obtiene el tipo subyacente y se permite DBNull.
            if (underlyingType == null)
            {
                underlyingType = dataType;
            }
            else
            {
                allowDBNull = true;
            }

            // Crea la columna con el tipo apropiado.
            DataColumn dataColumn = dt.Columns.Add(fieldName, underlyingType);
            dataColumn.AllowDBNull = allowDBNull;

            // Asigna el caption si se proporciona.
            if (!string.IsNullOrEmpty(captionColumn))
                dataColumn.Caption = captionColumn;

            // Marca la columna para exclusión si es necesario.
            if (excludeColumnFlag)
                dataColumn.ExtendedProperties[ExcludeColumnKey] = excludeColumnFlag;

            // Agrega la descripción si se proporciona.
            if (!string.IsNullOrEmpty(description))
                dt.AddDescription(dataColumn.ColumnName, description);

            return dataColumn;
        }

        /// <summary>
        /// Agrega o elimina una descripción para una columna existente.
        /// </summary>
        /// <param name="dt"><see cref="DataTable"/> que contiene la columna.</param>
        /// <param name="columnName">Nombre de la columna.</param>
        /// <param name="description">Descripción a asignar. Si es nula o vacía, se elimina la descripción existente.</param>
        public static void AddDescription(this DataTable dt, string columnName, string description)
        {
            if (dt == null || !dt.Columns.Contains(columnName))
                return;

            if (string.IsNullOrEmpty(description))
            {
                // Elimina la descripción si existe.
                if (dt.Columns[columnName].ExtendedProperties.ContainsKey(ColumnDescriptionKey))
                    dt.Columns[columnName].ExtendedProperties.Remove(ColumnDescriptionKey);
            }
            else
            {
                // Asigna la nueva descripción.
                dt.Columns[columnName].ExtendedProperties[ColumnDescriptionKey] = description;
            }
        }
        /// <summary>
        /// Obtiene la descripción almacenada para una columna específica.
        /// </summary>
        /// <param name="dt"><see cref="DataTable"/> que contiene la columna.</param>
        /// <param name="columnName">Nombre de la columna.</param>
        /// <returns>La descripción de la columna si existe; de lo contrario, <see cref="string.Empty"/>.</returns>
        public static string GetDescription(this DataTable dt, string columnName)
        {
            if (dt == null || !dt.Columns.Contains(columnName))
                return string.Empty;

            var column = dt.Columns[columnName];

            if (column.ExtendedProperties.ContainsKey(ColumnDescriptionKey))
            {
                // Obtiene el valor y lo convierte a string, manejando null
                var description = column.ExtendedProperties[ColumnDescriptionKey];
                return description?.ToString() ?? string.Empty;
            }

            return string.Empty;
        }

        /// <summary>
        /// Obtiene la descripción almacenada para una columna específica.
        /// </summary>
        /// <param name="column">La columna de la cual obtener la descripción.</param>
        /// <returns>La descripción de la columna si existe; de lo contrario, <see cref="string.Empty"/>.</returns>
        public static string GetDescription(this DataColumn column)
        {
            if (column == null)
                return string.Empty;

            if (column.ExtendedProperties.ContainsKey(ColumnDescriptionKey))
            {
                var description = column.ExtendedProperties[ColumnDescriptionKey];
                return description?.ToString() ?? string.Empty;
            }

            return string.Empty;
        }
    }
}