using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Boletos
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            using (var fldrDlg = new FolderBrowserDialog() { Description = "Selecione a pasta com os boletos:" })
            {
                fldrDlg.ShowNewFolderButton = false;

                if (fldrDlg.ShowDialog() == DialogResult.OK)
                {
                    string[] filePaths = Directory.GetFiles(fldrDlg.SelectedPath, "*.pdf", SearchOption.AllDirectories);

                    var listaBoletos = ExtrairBoletos(filePaths);

                    if (listaBoletos.Count > 0)
                    {
                        int maxlen = listaBoletos.Max(x => x.Valor.Length);
                        var excelFilePath = SaveExcelFile(listaBoletos.OrderBy(p => p.Eolica).ThenBy(c => c.Valor.PadLeft(maxlen, '0')));

                        if (File.Exists(excelFilePath))
                        {
                            Process.Start(excelFilePath);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Nenhum boleto encontrado na pasta selecionada.");
                    }
                }
            }
        }

        static List<Boleto> ExtrairBoletos(string[] pdfPaths)
        {
            var listaBoletos = new List<Boleto>();

            var listaCnpjs = new Dictionary<string, string>{
                { "12053787000210", "SANTA MARIA" },
                { "12053929000249", "SANTA HELENA" },
                { "14583703000102", "SANTO URIEL" },
                { "21216892000132", "SÃO BENTO DO NORTE I" },
                { "21216877000194", "SÃO BENTO DO NORTE II" },
                { "21216857000113", "SÃO BENTO DO NORTE III" },
                { "21216915000109", "SÃO MIGUEL I" },
                { "21216925000144", "SÃO MIGUEL II" },
                { "21216439000126", "SÃO MIGUEL III" },
                { "21957870000123", "GUAJIRU" },
                { "21957722000109", "JANGADA" },
                { "21957968000180", "POTIGUAR" },
                { "21917808000108", "CUTIA" },
                { "21909793000136", "MARIA HELENA" },
                { "21916951000185", "ESPERANÇA DO NORDESTE" },
                { "21909032000184", "PARAÍSO DOS VENTOS DO NORDESTE" },
                { "12723413000183", "BOA VISTA S/A" },
                { "12723335000117", "FAROL" },
                { "12723444000134", "OLHO D'ÁGUA" },
                { "12723384000150", "SÃO BENTO DO NORTE" },
                { "12802855000115", "ASA BRANCA I" },
                { "12802844000135", "ASA BRANCA II" },
                { "12802835000144", "ASA BRANCA III" },
                { "12802866000103", "NOVA EURUS IV" } 
                { "30097726000155", "VILA MARANHÃO I" }
                { "31004703000111", "VILA MARANHÃO II" }
                { "31449173000115", "VILA MARANHÃO III" }
                { "31478575000148", "VILA CEARÁ I" }
                { "34109229000180", "VILA MATO GROSSO I" }
                { "35823538000261", "JANDAIRA I" }
                { "35824347000214", "JANDAIRA II" }
                { "35823536000272", "JANDAIRA III" }
                { "35823577000269", "JANDAIRA IV" }};

            var listaEolicas = new List<string>(new string[] { "PARA(I|Í)SO.*NORDESTE", "VILA MARANH(Ã|A)O", "SANTA MARIA", "SANTA HELENA", "SANTO URIEL", "S(A|Ã)O BENTO DO NORTE I*", "S(A|Ã)O MIGUEL I*",
                "GUAJIRU", "JANGADA", "POTIGUAR", "CUTIA", "MARIA HELENA?", "ESPERAN(Ç|C)A DO NORDESTE", "BOA VISTA S/A", "FAROL", "OLHO D'?(Á|A)GUA", "ASA BRANCA I*", "NOVA EURUS IV" });

            foreach (var pdfPath in pdfPaths)
            {
                string text = string.Empty;

                try
                {
                    using (PdfReader reader = new PdfReader(pdfPath))
                    {
                        for (int page = 1; page <= reader.NumberOfPages; page++)
                        {
                            text += PdfTextExtractor.GetTextFromPage(reader, page);
                        }
                    }
                } catch(Exception e)
                {
                    Console.Error.WriteLine($"{pdfPath} - {e.Message}");
                    MessageBox.Show($"Não foi possível ler o pdf {pdfPath}", "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                var codigoBarrasMatch = Regex.Match(text, @".*(\d{5})[.\s]*(\d{5})[.\s]*(\d{5})[.\s]*(\d{6})[.\s]*(\d{5})[.\s]*(\d{6})[.\s]*(\d)[.\s]*(\d{14}).*", RegexOptions.Singleline);

                var textNumbers = Regex.Replace(text, "[^0-9]", "");

                var cnpjMatch = Regex.Match(textNumbers, string.Join("|", listaCnpjs.Keys));

                string eolica = string.Empty;

                if (!cnpjMatch.Success)
                {
                    var eolicaMatch = Regex.Match(text, string.Join("|", listaEolicas), RegexOptions.IgnoreCase);
                    eolica = eolicaMatch.Value;
                }

                var boleto = new Boleto();

                if (codigoBarrasMatch.Success)
                {
                    var codigoBarras = string.Join("", codigoBarrasMatch.Groups.Cast<System.Text.RegularExpressions.Group>().Skip(1).Select(o => o.Value));
                    var valorMatch = Regex.Match(codigoBarras.Substring(codigoBarras.Length - 10), @"^0+(?!$)(\d+)(\d{2})$");

                    if (valorMatch.Success)
                    {
                        var valor = $"{valorMatch.Groups[1]}.{valorMatch.Groups[2]}";
                        boleto.CodigoDeBarras = codigoBarras;
                        boleto.Valor = valor;

                        if (cnpjMatch.Success)
                        {
                            boleto.Cnpj = cnpjMatch.Value.Insert(2, ".").Insert(6, ".").Insert(10, "/").Insert(15, "-");

                            if (listaCnpjs.TryGetValue(cnpjMatch.Value, out var value))
                            {
                                boleto.Eolica = value;
                            }
                        }
                        else
                        {
                            boleto.Eolica = eolica;
                        }

                        listaBoletos.Add(boleto);
                    }
                }
            }

            return listaBoletos;
        }

        static string SaveExcelFile(IEnumerable<Boleto> listaBoletos)
        {
            var tempDir = System.IO.Path.GetTempPath();
            var filePath = System.IO.Path.Combine(tempDir, "Boletos.xlsx");

            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(filePath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Boletos"
            };
            sheets.Append(sheet);

            Cell nome = InsertCell(1, 0, worksheetPart.Worksheet);
            nome.CellValue = new CellValue("VALOR");
            nome.DataType = CellValues.String;

            Cell cargo = InsertCell(1, 1, worksheetPart.Worksheet);
            cargo.CellValue = new CellValue("CÓDIGO DE BARRAS");
            cargo.DataType = CellValues.String;

            Cell eolica = InsertCell(1, 2, worksheetPart.Worksheet);
            eolica.CellValue = new CellValue("CNPJ");
            eolica.DataType = CellValues.String;

            Cell cnpjCell = InsertCell(1, 3, worksheetPart.Worksheet);
            cnpjCell.CellValue = new CellValue("EÓLICA");
            cnpjCell.DataType = CellValues.String;

            uint count = 2;

            foreach (var boleto in listaBoletos)
            {
                nome = InsertCell(count, 0, worksheetPart.Worksheet);
                nome.CellValue = new CellValue(boleto.Valor);
                nome.DataType = CellValues.Number;

                cargo = InsertCell(count, 1, worksheetPart.Worksheet);
                cargo.CellValue = new CellValue(boleto.CodigoDeBarras);
                cargo.DataType = CellValues.String;

                cnpjCell = InsertCell(count, 2, worksheetPart.Worksheet);
                cnpjCell.CellValue = new CellValue(boleto.Cnpj);
                cnpjCell.DataType = CellValues.String;

                eolica = InsertCell(count, 3, worksheetPart.Worksheet);
                eolica.CellValue = new CellValue(boleto.Eolica);
                eolica.DataType = CellValues.String;

                count++;
            }

            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();

            return filePath;
        }

        static Cell InsertCell(uint rowIndex, uint columnIndex, Worksheet worksheet)
        {
            Row row = null;
            var sheetData = worksheet.GetFirstChild<SheetData>();

            // Check if the worksheet contains a row with the specified row index.
            row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // Convert column index to column name for cell reference.
            var columnName = GetColumnName((int)columnIndex);
            var cellReference = columnName + rowIndex;      // e.g. A1

            // Check if the row contains a cell with the specified column name.
            var cell = row.Elements<Cell>()
                       .FirstOrDefault(c => c.CellReference.Value == cellReference);
            if (cell == null)
            {
                cell = new Cell() { CellReference = cellReference };
                if (row.ChildElements.Count < columnIndex)
                    row.AppendChild(cell);
                else
                    row.InsertAt(cell, (int)columnIndex);
            }

            return cell;
        }

        static string GetColumnName(int index) // zero-based
        {
            const byte BASE = 'Z' - 'A' + 1;
            string name = string.Empty;

            do
            {
                name = Convert.ToChar('A' + index % BASE) + name;
                index = index / BASE - 1;
            }
            while (index >= 0);

            return name;
        }
    }
}
