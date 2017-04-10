using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Matstafett
{
    public class Shuffles
    {

        /// <summary>
        /// C# implementation of a Fisher-Yates Shuffle
        /// </summary>
        /// <param name="length">the length of the array to create.</param>
        /// <returns>shuffled array of integers</returns>
        public int[] FisherYatesShuffleArray(int length)
        {
            int[] array = Enumerable.Range(0, length).ToArray();
            Random rand = new Random();
            for (int i = length - 1; i > 0; i--)
            {
                int n = rand.Next(i + 1);
                int temp = array[i];
                array[i] = array[n];
                array[n] = temp;
            }
            return array;
        }
    }
}