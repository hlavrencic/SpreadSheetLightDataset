using System.Collections.Generic;
using Tools.SpreadSheets.Data;

namespace Tools.SpreadSheets
{
    /// <summary>
    /// Se encarga crear el archivo
    /// </summary>
    public interface ISpreadSheetGenerator
    {
        /// <summary>
        /// Crea un nuevo archivo a partir de un template.
        /// Y le escribe los datos recibidos por parametro
        /// </summary>
        /// <param name="pathTemplate">Ruta al archivo template</param>
        /// <param name="newFilePath">Ruta donde se guardara el archivo nuevo generado</param>
        /// <param name="dataAliasWriters">Lista de datos</param>
        void CreateFromTemplate(
            string pathTemplate,
            string newFilePath,
            params IDataAliasWriter[] dataAliasWriters);

        /// <summary>
        /// Crea un nuevo archivo a partir de un template.
        /// Y le escribe los datos recibidos por parametro
        /// </summary>
        /// <param name="pathTemplate">Ruta al archivo template</param>
        /// <param name="newFilePath">Ruta donde se guardara el archivo nuevo generado</param>
        /// <param name="dataAliasWriters">Lista de datos</param>
        void CreateFromTemplate(
            string pathTemplate,
            string newFilePath,
            string sheetName,
            params IDataAliasWriter[] dataAliasWriters
            );
    }
}
