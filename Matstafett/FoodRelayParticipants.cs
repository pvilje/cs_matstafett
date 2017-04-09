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

        public void AddParticipant(Participant newParticipant)
        {
            this.All.Add(newParticipant);
        }

        /* ValidateNumberOfParticipants
         * Verifies the number of found participants. 
         * Returns: 0 - OK, 1 - too few, 2 - not a factor of three.
         */
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

        public void GenerateRandomizedIndex()
        {
            this.RandomizedIndex = Enumerable.Range(0, this.NumberOfParticipants + 1).ToArray();
        }
    }
}
