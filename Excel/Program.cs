using System.Text;
using Excel;
using OfficeOpenXml;


namespace ExcelConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Excel file here
            var currentDirectory = Directory.GetCurrentDirectory();
            var excelFiles = Directory.GetFiles(currentDirectory, "*.xlsx");

            if (!excelFiles.Any())
            {
                Console.WriteLine("No Excel files found in the current directory.");
                return; // Exit the program if no Excel files are found
            }

            var filePath = excelFiles[0];
            // var filePath

            // Ensure the license context is set (for EPPlus 5 and above)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];  // Gets the first worksheet

            int rowCount = worksheet.Dimension.Rows;
            // int colCount = worksheet.Dimension.Columns;

            // Header count
            int nonEmptyHeaderCount = 1;
            while (!string.IsNullOrEmpty(worksheet.Cells[1, nonEmptyHeaderCount].Text))
            {
                nonEmptyHeaderCount++;
            }
            int colCount = nonEmptyHeaderCount - 1;

            List<CusCol> colRrep = new List<CusCol>(); //Columns info

            // Loop through the columns
            for (int col = 1; col <= colCount; col++)
            {
                string columnName = worksheet.Cells[1, col].Text;

                // Get column name and length
                int nonEmptyCellCount = 0;
                for (int row = 2; row <= rowCount; row++)  // Starting from 2nd row since 1st row is header
                {
                    if (!string.IsNullOrEmpty(worksheet.Cells[row, col].Text))
                    {
                        nonEmptyCellCount++;
                    }
                }

                // int re = (int)Char.GetNumericValue(columnName[0]); // Get Repeatability
                int re = Int32.Parse(columnName);

                colRrep.Add(new CusCol { Cnum = col, Cname = columnName, Clength = nonEmptyCellCount, Reable = re });
            }

            var outputsheet = package.Workbook.Worksheets[1];  // Gets the Output number
            int outputNum;
            if (string.IsNullOrWhiteSpace(outputsheet.Cells[1, 1].Text))
            {
                outputNum = 10;
            }
            else
            {
                outputNum = Convert.ToInt32(outputsheet.Cells[1, 1].Text);
            }

            List<List<string>> ranCol = new List<List<string>>();
            Random random = new Random();
            for (int i = 0; i < outputNum; i++) // Yeild
            {
                List<string> data = new();
                foreach (CusCol c in colRrep)
                {

                    int ran1;
                    if (c != null)
                    {
                        ran1 = random.Next(2, c.Clength + 1);
                    }
                    else
                    {
                        throw new NullReferenceException("Column is null!");
                    }
                    // Console.WriteLine(c.Clength + "  "+ ran1);

                    string rep1 = worksheet.Cells[ran1, c.Cnum].Text;
                    string rep2 = worksheet.Cells[ran1, c.Cnum].Text;
                    if (c.Reable != 1 && random.Next(0, 2) == 1)
                    {
                        int ran2 = random.Next(2, c.Clength + 1);
                        while (ran1 == ran2)
                        {
                            ran2 = random.Next(2, c.Clength + 1);
                        }
                        rep2 = worksheet.Cells[ran2, c.Cnum].Text;
                    }

                    //     if (rep1 == rep2)
                    //     {
                    //         data.Add(rep1);
                    //         Console.Write(c.Cnum + "-" + rep1 + " ");  // Show Test
                    //     }
                    //     else
                    //     {
                    //         data.Add(rep1);
                    //         data.Add(rep2);
                    //         Console.Write(c.Cnum + "-" + rep1 + " ");
                    //         Console.Write(c.Cnum + "-" + rep2 + " ");
                    //     }
                    // }

                    if (rep1 == rep2)
                    {
                        data.Add(rep1);
                        Console.Write(rep1 + " "); // Show Output
                    }
                    else
                    {
                        data.Add(rep1);
                        data.Add(rep2);
                        Console.Write(rep1 + " ");
                        Console.Write(rep2 + " ");
                    }
                }
                ranCol.Add(data);
                Console.WriteLine();
            }

            static string AssembleText(List<List<string>> data)
            {
                StringBuilder sb = new StringBuilder();
                foreach (var row in data)
                {
                    sb.AppendLine(string.Join(" ", row));
                }
                return sb.ToString();
            }

            string outputText = AssembleText(ranCol);
            string outputFilePath = Path.ChangeExtension(filePath, "Output.txt");  // Change the extension to .txt

            File.WriteAllText(outputFilePath, outputText);

            Console.WriteLine($"Data written to {outputFilePath}");

            // Display the row names and their counts
            // foreach (var entry in colRrep)
            // {
            //     // Console.WriteLine($"Cnum: '{entry.Cnum}' Cname: '{entry.Cname}' Clength: '{entry.Clength}' CRe: '{entry.Reable}'");
            //     // string combinedString = string.Join( " ", ranCol);
            //     // Console.WriteLine(combinedString);
            // }
        }
    }
}
