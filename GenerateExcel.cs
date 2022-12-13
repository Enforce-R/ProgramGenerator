
using ClosedXML.Excel;
using ProgramGenerator.Models;

namespace ProgramGenerator
{
    public class GenerateExcel
    {
        public static async Task CreateExcelFile(List<Part> parts, string title)
        {
            try
            {
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add(title);
                worksheet.Columns().AdjustToContents(1);
                worksheet.Rows().AdjustToContents();

                worksheet.Cell("A1").Value = "Εβδομάδα";
                worksheet.Cell("A1").Style.Font.SetFontColor(XLColor.White).Fill.SetBackgroundColor(XLColor.Black);
                worksheet.Cell("B1").Value = "Εισαγωγή";
                worksheet.Cell("B1").Style.Font.SetFontColor(XLColor.White).Fill.SetBackgroundColor(XLColor.Black);

                worksheet.Cell("C1").Value = "Προσευχή";

                worksheet.Cell("E1").Value = "Ομιλία";
                worksheet.Cell("E2").Value = "Θέμα";
                worksheet.Cell("D2").Value = "Διάρκεια";

                worksheet.Cell("D2").Value = "Διάρκεια";

                worksheet.Cell("H1").Value = "Πετράδια";
                worksheet.Cell("H2").Value = "Θέμα";
                worksheet.Cell("G2").Value = "Διάρκεια";

                worksheet.Cell("K1").Value = "Ανάγνωση";
                worksheet.Cell("J2").Value = "Διάρκεια";
                worksheet.Cell("K2").Value = "Πηγή";
                worksheet.Cell("L2").Value = "Μέρος";

                worksheet.Cell("O1").Value = "Διακονία 1";
                worksheet.Cell("N2").Value = "Διάρκεια";
                worksheet.Cell("O2").Value = "Πηγή";
                worksheet.Cell("P2").Value = "Μέρος";

                worksheet.Cell("S1").Value = "Διακονία 2";
                worksheet.Cell("R2").Value = "Διάρκεια";
                worksheet.Cell("S2").Value = "Πηγή";
                worksheet.Cell("T2").Value = "Μέρος";

                worksheet.Cell("W1").Value = "Διακονία 3";
                worksheet.Cell("V2").Value = "Διάρκεια";
                worksheet.Cell("W2").Value = "Πηγή";
                worksheet.Cell("X2").Value = "Μέρος";

                worksheet.Cell("AA1").Value = "Χριστιανοί 1";
                worksheet.Cell("Z2").Value = "Διάρκεια";
                worksheet.Cell("AA2").Value = "Πηγή";
                worksheet.Cell("AB2").Value = "Μέρος";

                worksheet.Cell("AE1").Value = "Χριστιανοί 2";
                worksheet.Cell("AD2").Value = "Διάρκεια";
                worksheet.Cell("AE2").Value = "Πηγή";
                worksheet.Cell("AF2").Value = "Μέρος";

                worksheet.Cell("AI1").Value = "Χριστιανοί 3";
                worksheet.Cell("AH2").Value = "Διάρκεια";
                worksheet.Cell("AI2").Value = "Πηγή";
                worksheet.Cell("AJ2").Value = "Μέρος";

                worksheet.Cell("AM1").Value = "Μελέτη";
                worksheet.Cell("AL2").Value = "Διάρκεια";
                worksheet.Cell("AM2").Value = "Πηγή";
                worksheet.Cell("AN2").Value = "Μέρος";

                worksheet.Cell("AQ1").Value = "Προσευχή";
                worksheet.Cell("AR1").Value = "Ύμνος";

                for (var i = 0; i < parts.Count; i++)
                {
                    Part part = parts[i];
                    string partDescription = part.getPartDescription();
                    string partDuration = part.getPartDuration();

                    if (part.getPartType() == PartType.Treasures)
                    {
                        worksheet.Cell($"A{i + 2}").Style.Font.SetFontColor(XLColor.White).Fill.SetBackgroundColor(XLColor.Blue);
                    }
                    else if (part.getPartType() == PartType.Ministry)
                    {
                        worksheet.Cell($"A{i + 2}").Style.Font.SetFontColor(XLColor.White).Fill.SetBackgroundColor(XLColor.Black);
                    }
                    else if (part.getPartType() == PartType.Christians)
                    {
                        worksheet.Cell($"A{i + 2}").Style.Font.SetFontColor(XLColor.White).Fill.SetBackgroundColor(XLColor.Red);
                    }

                    worksheet.Cell($"A{i + 2}").Value = partDescription;
                    if (partDuration != "")
                    {
                        worksheet.Cell($"B{i + 2}").Value = part.getPartDuration() + " λεπτά";
                    }

                }
                workbook.SaveAs(@$"Ζωή και Διακονία\{title}.xlsx");
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }
}
