using ProgramGenerator.Models;
using ProgramGenerator;
using System.Reflection;
using System.Text.RegularExpressions;

namespace CSharpTutorials
{
    class Program
    {
        public static async Task Main(string[] args)
        {
            await TransformFileAsync();
        }

        static async Task TransformFileAsync()
        {
            try
            {
                string currentDirectory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string fileDirectory = @"Ζωή και Διακονία";
                if (!Directory.Exists(fileDirectory))
                {
                    Directory.CreateDirectory(fileDirectory);
                }
                string[] files = Directory.GetFiles(fileDirectory, "*.epub");

                List<Part> allParts = new();

                foreach (string file in files)
                {
                    List<Part> myData = await DataHandling.ExtractData(file);
                    allParts.AddRange(myData);
                }
                await GenerateExcel.CreateExcelFile(allParts, "Generated_Πρόγραμμα");
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

// Ανάγνωση να έχει ύλη. 
// 