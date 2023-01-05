using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProgramGenerator.Models
{
    public class Part
    {
        private string week { get; set; }
        private PartType typeOfPart { get; set; }
        private string duration { get; set; }
        private string description { get; set; }

        public Part(string week, PartType typeOfPart, string duration, string description)
        {
            this.week = week;
            this.typeOfPart = typeOfPart;
            this.duration = duration;
            this.description = description;
        }

        public string getPartWeek()
        {
            return week;
        }
        public PartType getPartType()
        {
            return typeOfPart;
        }

        public string getPartDuration()
        {
            return duration;
        }

        public string getPartDescription()
        {
            return description;
        }

    }

    public enum PartType
    {
        Verse,
        Week,
        Presenter,
        Song,
        Prayer,
        Treasures,
        TreasureSpeech,
        TreasureGems,
        BibleReading,
        Ministry,
        MinistryPart,
        Christians,
        ChristiansPart,
        BibleStudy
    }
}
