using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Matstafett
{
    public class FoodRelayParticipants
    {
        public List<Participant> AllStarterHosts { get; set; }
        public List<Participant> AllMaincourseHosts { get; set; }
        public List<Participant> AllDesertHosts { get; set; }
        public List<Participant> All { get; set; }
        public List<Participant> AllSorted { get; set; }
        public List<Participant> FinalStarterHosts { get; set; }
        public List<Participant> FinalStarterGuests1 { get; set; }
        public List<Participant> FinalStarterGuests2 { get; set; }
        public List<Participant> FinalMainCourseHosts { get; set; }
        public List<Participant> FinalMainCourseGuests1 { get; set; }
        public List<Participant> FinalMainCourseGuests2 { get; set; }
        public List<Participant> FinalDesertHosts { get; set; }
        public List<Participant> FinalDesertGuests1 { get; set; }
        public List<Participant> FinalDesertGuests2 { get; set; }
        public int ParticipantsPerGroup { get; private set; }
        public int NumberOfParticipants { get; private set; }
        public int[] RandomizedIndex { get; private set; }

        public FoodRelayParticipants()
        {
            this.AllStarterHosts = new List<Participant>();
            this.AllMaincourseHosts = new List<Participant>();
            this.AllDesertHosts = new List<Participant>();
            this.All = new List<Participant>();
            this.AllSorted = new List<Participant>();
            this.FinalStarterHosts = new List<Participant>();
            this.FinalStarterGuests1 = new List<Participant>();
            this.FinalStarterGuests2 = new List<Participant>();
            this.FinalMainCourseHosts = new List<Participant>();
            this.FinalMainCourseGuests1 = new List<Participant>();
            this.FinalMainCourseGuests2 = new List<Participant>();
            this.FinalDesertHosts = new List<Participant>();
            this.FinalDesertGuests1 = new List<Participant>();
            this.FinalDesertGuests2 = new List<Participant>();
            this.NumberOfParticipants = 0;
        }

        /// <summary>
        /// Adds a new participant to the list 
        /// </summary>
        /// <param name="newParticipant">The new participant</param>
        public void AddParticipant(Participant newParticipant)
        {
            this.All.Add(newParticipant);
        }

        /// <summary>
        /// Verifies the number of found participants. 
        /// </summary>
        /// <returns>0 - OK, 1 - too few, 2 - not a factor of three.</returns>
        public int ValidateNumberOfParticipants()
        {
            if (this.All.Count < 9)
            {
                return 1;
            }
            else if (this.All.Count % 3 != 0)
            {
                return 2;
            }
            this.ParticipantsPerGroup = this.All.Count / 3;
            this.NumberOfParticipants = this.All.Count;
            return 0;
        }

        /// <summary>
        /// Generates a ranodimized list of integers 
        /// of the same length as the number of participants
        /// </summary>
        public void GenerateRandomizedIndex()
        {
            this.RandomizedIndex = new Shuffles().FisherYatesShuffleArray(this.NumberOfParticipants);
        }

        /// <summary>
        /// Place participants into their groups.
        /// </summary>
        public void PlaceParticipantsIntoGroups()
        {
            foreach(int index in RandomizedIndex)
            {
                AllSorted.Add(All[index]);
            }
            for (int unorderIndex = 0; unorderIndex < this.NumberOfParticipants; unorderIndex++)
            {
                if (unorderIndex < this.ParticipantsPerGroup)
                {
                    this.AllStarterHosts.Add(this.AllSorted[this.RandomizedIndex[unorderIndex]]);
                }
                else if (unorderIndex < this.ParticipantsPerGroup * 2)
                {
                    this.AllMaincourseHosts.Add(this.AllSorted[RandomizedIndex[unorderIndex]]);
                }
                else
                {
                    this.AllDesertHosts.Add(this.AllSorted[RandomizedIndex[unorderIndex]]);
                }
            }
        }

        /// <summary>
        /// Create the final lineup.
        /// </summary>
        public void GenerateLineup()
        {
            int baseIndex = 0;
            int offset1 = 1;
            int offset2 = 2;
            while(baseIndex < ParticipantsPerGroup)
            {
                offset1 = (offset1 >= ParticipantsPerGroup) ? 0 : offset1;
                offset2 = (offset2 >= ParticipantsPerGroup) ? 0 : offset2;

                // Starters
                FinalStarterHosts.Add(AllStarterHosts[baseIndex]);
                FinalStarterGuests1.Add(AllMaincourseHosts[baseIndex]);
                FinalStarterGuests2.Add(AllDesertHosts[baseIndex]);

                // Main Course
                FinalMainCourseHosts.Add(AllMaincourseHosts[offset1]);
                FinalMainCourseGuests1.Add(AllStarterHosts[baseIndex]);
                FinalMainCourseGuests2.Add(AllDesertHosts[offset2]);

                // Desert
                FinalDesertHosts.Add(AllDesertHosts[offset1]);
                FinalDesertGuests1.Add(AllStarterHosts[baseIndex]);
                FinalDesertGuests2.Add(AllMaincourseHosts[offset2]);

                baseIndex++;
                offset1++;
                offset2++;
            }
        }
    }
}
