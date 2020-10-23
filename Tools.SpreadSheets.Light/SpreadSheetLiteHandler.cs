using Tools.SpreadSheets.Data;
using SpreadsheetLight;
using Tools.SpreadSheets.Handlers;
using System;
using System.Linq;

namespace Tools.SpreadSheets.Light
{
    public class SpreadSheetLiteHandler : ISpreadSheetHandler
    {
        private readonly SLDocument sLDocument;

        public SpreadSheetLiteHandler(
            SLDocument sLDocument)
        {
            this.sLDocument = sLDocument;
        }

        public void Write<TData>(TData data, int rowIndex, int colIndex)
        {
            if(data == null)
            {
                sLDocument.SetCellValue(rowIndex, colIndex, string.Empty);
                return;
            }

            var dataType = data.GetType();

            if (data.GetType() == typeof(Int32))
            {
                int dataInt = Int32.Parse(data.ToString());
                sLDocument.SetCellValue(rowIndex, colIndex, dataInt);
            }
            else if (data.GetType() == typeof(float))
            {
                float dataFloat = float.Parse(data.ToString());
                sLDocument.SetCellValue(rowIndex, colIndex, dataFloat);
            }
            else if (data.GetType() == typeof(double))
            {
                double dataDouble = double.Parse(data.ToString());
                sLDocument.SetCellValue(rowIndex, colIndex, dataDouble);
            }
            else if (data.GetType() == typeof(decimal))
            {
                decimal dataDecimal = decimal.Parse(data.ToString());
                sLDocument.SetCellValue(rowIndex, colIndex, dataDecimal);
            }
            else if (data.GetType() == typeof(DateTime))
            {
                DateTime dataDateTime = DateTime.Parse(data.ToString());
                sLDocument.SetCellValue(rowIndex, colIndex, dataDateTime);
            }
            else if (data.GetType() == typeof(string))
            {
                string dataString = data.ToString();
                sLDocument.SetCellValue(rowIndex, colIndex, dataString);
            }
            else
            {
                var msgError = string.Format("Tipo no soportado {0}", dataType);
                throw new NotSupportedException(msgError);
            }
        }

        public void CopyRow(int rowSourceIndex)
        {
            var newRowIndex = rowSourceIndex + 1;
            sLDocument.InsertRow(newRowIndex, 1);
            sLDocument.CopyRow(rowSourceIndex, newRowIndex);
        }

        public void CopyCol(int colSourceIndex)
        {
            var newColIndex = colSourceIndex + 1;
            sLDocument.InsertColumn(newColIndex, 1);
            sLDocument.CopyColumn(colSourceIndex, newColIndex);
        }

        public void ShiftRight(int rowIndex, int colIndex)
        {
            
            //sLDocument.get
            //sLDocument.CopyCell()
            throw new NotImplementedException();
        }

        /// <summary>
        /// Devuelve el dato con las coordenadas cargadas. 
        /// Devuelve Nulo si no encuentra la ubicación.
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public IDataCoordWriter GetPosition(IDataAliasWriter data)
        {
            var sheetCoords = sLDocument.GetDefinedNameText(data.Alias);
            var sheetName = String.IsNullOrEmpty(sheetCoords) ? null : sheetCoords.Split('!')[0];            
            var colRow = String.IsNullOrEmpty(sheetCoords) ? null : sheetCoords.Split('!')[1].Split('$').Where(x => !string.IsNullOrEmpty(x)).ToArray();

            if (sheetName == null && colRow == null)
                return null;

            //TODO[Mejora] Utilizar esta funcion --> SLDocument.WhatIsRowColumnIndex
            return new DataCoordWriter(data)
            {
                ColIndex = char.ToUpper(colRow[0][0]) - 64,
                RowIndex = Int32.Parse(colRow[1].ToString()),
                SheetName = sheetName
            };
        }

        public bool SelectWorksheet(string workSheetName)
        {
            return sLDocument.SelectWorksheet(workSheetName);
        }

        public void Save(string path)
        {
            sLDocument.SaveAs(path);
        }

        public void Dispose()
        {
            sLDocument.Dispose();
        }

        public void RenameWorkSheet(string name)
        {
            // The first worksheet of a new spreadsheet is named "Sheet1",
            // but use the constant to future-proof your application.
            sLDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, name);
        }
    }
}
