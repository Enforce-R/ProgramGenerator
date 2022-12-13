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

                foreach (string file in files)
                {
                    List<Part> myData = await DataHandling.ExtractData(file);
                    string dates = myData[0].getPartWeek() + "_" + myData[myData.Count - 1].getPartWeek();
                    string title = Regex.Replace(dates, "[^α-ωΑ-Ωίέώύόά_]", "");
                    await GenerateExcel.CreateExcelFile(myData, title);
                }
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