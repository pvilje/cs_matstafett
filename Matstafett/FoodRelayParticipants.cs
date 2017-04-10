using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Matstafett
{
    public class FoodRelayParticipants
    {
        public List<Participant> StarterHosts { get; set; }
        public List<Participant> MaincourseHosts { get; set; }
        public List<Participant> DesertHosts { get; set; }
        public List<Participant> All { get; set; }
        public int ParticipantsPerGroup { get; private set; }
        public int NumberOfParticipants { get; private set; }
        public int[] RandomizedIndex { get; private set; }

        public FoodRelayParticipants()
        {
            this.StarterHosts = new List<Participant>();
            this.MaincourseHosts = new List<Participant>();
            this.DesertHosts = new List<Participant>();
            this.All = new List<Participant>();
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
            Participant[] unorderedParticipantArray = new Participant[this.NumberOfParticipants];
            for (int unorderIndex = 0; unorderIndex < this.NumberOfParticipants; unorderIndex++)
            {
                if (unorderIndex < this.ParticipantsPerGroup)
                {
                    this.StarterHosts.Add(this.All[this.RandomizedIndex[unorderIndex]]);
                }
                else if (unorderIndex < this.ParticipantsPerGroup * 2)
                {
                    this.MaincourseHosts.Add(this.All[RandomizedIndex[unorderIndex]]);
                }
                else
                {
                    this.DesertHosts.Add(this.All[RandomizedIndex[unorderIndex]]);
                }
            }
        }
    }
}
