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

        public FoodRelayParticipants()
        {
            this.StarterHosts = new List<Participant>();
            this.MaincourseHosts = new List<Participant>();
            this.DesertHosts = new List<Participant>();
            this.All = new List<Participant>();
        }
        public void AddParticipant(Participant newParticipant)
        {
            this.All.Add(newParticipant);
        }
    }
}
