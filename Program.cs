using System;
using System.IO;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using NPOI.XSSF.UserModel;

namespace XMLToExcel
{
    class Program
    {
        static List<string> CombineBasedOnString(IEnumerable<string> input, List<string> tokens)
        {
            var listInput = input.ToList();
            var output = new List<string>();
            string elt = "";
            //var rowGroup = false;
            for (int i = 0; i < listInput.Count; i++)
            {
                var line = listInput[i];
                elt += (line + "\n");
                if (line == tokens[0])
                {
                    for (int j = 0; j < tokens.Count; j++)
                    {
                        var token = tokens[j];
                        if (i + 1 < listInput.Count)
                        {
                            if (j + 1 < tokens.Count)
                            {
                                //if (listInput[i].Contains("</Row>") && listInput[i + 1].Contains("</RowGroup>"))
                                //    rowGroup = true;
                                if (listInput[i + 1] != tokens[j + 1])
                                {
                                    output.Add(elt);
                                    elt = "";
                                }
                                else
                                {
                                    line = listInput[++i];
                                    elt += (line + "\n");                                        
                                    if (i == listInput.Count - 1)
                                    {
                                        output.Add(elt);
                                        break;
                                    }
                                    continue;
                                }
                            }
                            else
                            {
                                if (line == token)
                                {
                                    output.Add(elt);
                                    elt = "";
                                }
                            }
                        }
                        else
                        {
                            elt += listInput[i + 1];
                            output.Add(elt);
                            break;
                        }
                    }
                }
            }
            return output;
        }
        static void Main(string[] args)
        {
            Console.WriteLine();
            var outputFile = @"D:\Svatantra\SheetKraftWeb\Templates\LiabilityMaturity.xls";
            var characterCount = 25000;
            var inputFile = @"C:\Users\QP2020\Downloads\liabilityMaturity.xml";
            string[] lines = File.ReadAllLines(inputFile);
            var groupedLines = CombineBasedOnString(lines, new List<string> { "\t\t\t\t</Row>", "\t\t\t</Table>", "\t\t</TableRow>", "\t</Worksheet>", "</Workbook>" });

            HSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(outputFile, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
                file.Close();
            }

            ISheet sheet = hssfwb.GetSheetAt(0);
            int j = 0;
            IRow row = sheet.GetRow(j);
            ICell cell = row.GetCell(0);
            int count = 0;
            string output = "";
            foreach (var line in groupedLines)
            {
                if (!String.IsNullOrEmpty(line))
                    output += ("\n" + line);
                else
                    continue;
                count += line.Length;
                if (count >= characterCount)
                {
                    cell.SetCellValue(output);
                    output = "";
                    row = sheet.GetRow(++j) ?? sheet.CreateRow(j);
                    cell = row.GetCell(0) ?? row.CreateCell(0);
                    count = 0;
                }
            }
            if (count > 0)
            {
                cell.SetCellValue(output);
            }
            using (FileStream file = new FileStream(outputFile, FileMode.Open, FileAccess.Write))
            {
                hssfwb.Write(file);
                file.Close();
            }
        }
    }
}
