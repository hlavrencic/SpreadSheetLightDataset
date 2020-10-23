using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NUnit.Framework;
using Tools.SpreadSheets.Data;
using Tools.SpreadSheets.Data.Helpers;
using Tools.SpreadSheets.Light;

namespace Tools.SpreadSheets.Tests
{
    [TestFixture]
    public class SpreadSheetLightTest
    {
        private string newFilePath;
        private ISpreadSheetGenerator spreadSheetGenerator;
        private string template;

        [TestFixtureSetUp]
        public void TestFixtureSetUp()
        {
            var pathString = TestContext.CurrentContext.TestDirectory;
            newFilePath = Path.Combine(pathString, "informe2.xlsx");
            template = Path.Combine(pathString, "informe.xlsx");
            spreadSheetGenerator = SpreadSheetLightHandlerFactory.CreateGenerator();
        }

        [TestCase]
        public void SpreadSheetLightTestIntegrador()
        {
            var dataSource1 = new[] 
            {
                new RegistroEjemplo { Id = 1, Nombre = "Nombre1" , Fecha = new DateTime(2000,1,1) },
                new RegistroEjemplo { Id = 2, Nombre = "Nombre2" , Fecha = new DateTime(2000,1,2) },
                new RegistroEjemplo { Id = 3, Nombre = "Nombre3" , Fecha = new DateTime(2000,1,3) },
            };

            var cell1 = dataSource1.First().Id.ToCell("cell1");
            var cell3 = dataSource1.First().Id.ToCell("cell3");

            var rowAlias1 = RegistroEjemploMapper(dataSource1.First()).ToRowData("rowAlias1");
            var rowAlias2 = RegistroEjemploMapper(dataSource1.First()).ToRowData("rowAlias2");

            // Cabecera
            var cabecera1 = RegistroEjemploMapper(dataSource1.First()).ToRowData("cabecera1");

            // Filas
            var filas = dataSource1.Select(RegistroEjemploMapper);

            var filasDinamicas = RegistroColDynamicEjemploMapper(filas);

            var range1 = filas.Select(s => s.ToRowData()).ToDynamicRangeData("range1");

            // Columnas dinamicas
            var range1dinamico = filasDinamicas.ToRangeData("range1dinamico");

            // Agregar rangeCol1 y rangeCol1dinamico
             spreadSheetGenerator.CreateFromTemplate(template, newFilePath, cell1, cell3, rowAlias1, rowAlias2, range1, range1dinamico);
        }

        private IEnumerable<IRowData> RegistroColDynamicEjemploMapper(IEnumerable<IEnumerable<IDataWriter>> data)
        {
            var cont = 0;
            foreach(var reg in data)
            { 
                if(cont == 0)
                {
                    var cabecera = new[] { "Id", "Nombre", "Fecha" }.Select(c => c.ToCell()).ToDynamicRowData();
                    yield return cabecera;
                    yield return reg.ToRowData();
                } 
                else if(cont < 2)
                {
                    yield return reg.ToRowCopyData();
                } else
                {
                    yield return reg.ToRowData();
                }

                cont++;
            }

        }

        private IEnumerable<IDataWriter> RegistroEjemploMapper(RegistroEjemplo data)
        {
            return new IDataWriter[]
            {
                data.Id.ToCell(),
                data.Nombre.ToCell(),
                data.Fecha.ToCell()
            };

            /*
            yield return data.Id.ToCell();
            yield return data.Nombre.ToCell();
            yield return data.Fecha.ToCell();
            */
        }

        public class RegistroEjemplo
        {
            public int Id { get; set; }

            public string Nombre { get; set; }

            public DateTime Fecha { get; set; }
        }
    }
}
