using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Matstafett
{
    public class Participant
    {
        public string Name { get; set; }
        public string ContactInformation { get; set; }
        public string Allergie { get; set; }

        public Participant(string name = null, string contact = null, string allergie = null )
        {
            // Default to " " if no value is submitted
            this.Name = name ?? " ";
            this.ContactInformation = contact ?? " ";
            this.Allergie = allergie ?? " ";
        }
    }
}
