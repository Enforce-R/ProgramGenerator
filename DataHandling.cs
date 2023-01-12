using EpubSharp;
using ProgramGenerator.Models;
using System.Text;
using System.Text.RegularExpressions;

namespace ProgramGenerator
{
    public class DataHandling
    {
        public static Task<List<Part>> ExtractData(string path)
        {
            EpubBook book = EpubReader.Read(path);

            string title = book.Title;
            ICollection<EpubChapter> chapters = book.TableOfContents;
            string text = book.ToPlainText().Substring(book.ToPlainText().IndexOf('©') + 1);

            byte[] bytes = Encoding.Default.GetBytes(text);
            text = Encoding.UTF8.GetString(bytes);
            //convert to a UTF

            List<string> weeks = new List<string>();
            string chapter = "";
            for (int i = 0; i < chapters.Count; i++)
            {
                chapter = chapters.ToList()[i].Title;
                bool isDate = Regex.IsMatch(chapter, "^\\d");

                if (isDate)
                {
                    weeks.Add(chapter);
                }
            }

            string weekContent = "";
            List<Part> parts = new List<Part>();

            for (int i = 0; i <= weeks.Count - 1; i++)
            {
                if (i == weeks.Count - 1)
                {
                    weekContent = getBetween(text, weeks[i], "^");
                    ContentTidyUp(weeks[i], weekContent, parts);
                    break;
                }
                else weekContent = getBetween(text, weeks[i], weeks[i + 1]);

                parts = ContentTidyUp(weeks[i], weekContent, parts);
            }

            return Task.FromResult(parts);
        }

        private static List<Part> ContentTidyUp(string week, string content, List<Part> parts)
        {
            string verse = "";
            using (var reader = new StringReader(content))
            {
                for (int i = 0; i < 2; i++)
                {
                    string line = reader.ReadLine();
                    if (i == 1)
                    {
                        verse = line;
                    }
                }
            }
            parts.Add(new Part(week, PartType.Week, "", week));
            parts.Add(new Part(week, PartType.Song, "", Regex.Match(getBetween(content, "\n\n\n \n ", "\n"), @"\d+").Value ));
            parts.Add(new Part(week, PartType.Presenter, "1", "Εισαγωγή"));
            parts.Add(new Part(week, PartType.TreasureSpeech, "10", "«" + getBetween(content, "«", "»") + "»"));
            parts.Add(new Part(week, PartType.TreasureGems, "10", verse));
            parts.Add(new Part(week, PartType.BibleReading, "4", getBetween(content, "Ανάγνωση της Αγίας Γραφής: (4 λεπτά)", "\n")));


            string contentMinistry = getBetween(content, "ΔΙΑΚΟΝΙΑ ΑΓΡΟΥ \n \n", " \n\n\n ΠΩΣ ΝΑ ΖΟΥΜΕ");
            string contentChristians = getBetween(content, "ΩΣ ΧΡΙΣΤΙΑΝΟΙ \n \n", "\n\n \n\n");
            using (var reader = new StringReader(contentMinistry))
            {
                for (string line = reader.ReadLine(); line != null; line = reader.ReadLine())
                {
                    if (line.Contains("th"))
                    {
                        parts.Add(new Part(week, PartType.MinistryPart, getBetween(line, ": (", "λεπτά"), line));
                    }
                    parts.Add(new Part(week, PartType.MinistryPart, getBetween(line, ": (", "λεπτά"), line));
                    // getBetween(line, line.Substring(0, 1), ": (") + line[(line.LastIndexOf('('))..])
                }
            }
            using (var reader = new StringReader(contentChristians))
            {
                for (string line = reader.ReadLine(); line != null; line = reader.ReadLine())
                {
                    if (line.Contains("Ύμνος"))
                    {
                        parts.Add(new Part(week, PartType.Song, "", Regex.Match(line, @"\d+").Value));
                    }
                    else if (line.Contains("Τελικά Σχόλια"))
                    {

                    }
                    else if (line.Contains("Εκκλησιαστική Γραφική Μελέτη"))
                    {
                        string timePart = getBetween(line, ": (", "λεπτά");
                        parts.Add(new Part(week, PartType.BibleStudy, timePart, line.Replace("Εκκλησιαστική Γραφική Μελέτη: (" + timePart + "λεπτά)", "")));
                    }
                    else
                    {
                        if(line != "")
                        {
                            parts.Add(new Part(week, PartType.ChristiansPart, getBetween(line, ": (", "λεπτά"), getBetween(line, line.Substring(0, 1), ": (")));
                        }
                    }
                }
            }
            return parts;
        }

        public static string getBetween(string strSource, string strStart, string strEnd)
        {
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                int Start, End;
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }

            return "";
        }
    }
}
