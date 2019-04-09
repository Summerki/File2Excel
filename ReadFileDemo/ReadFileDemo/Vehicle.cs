using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadFileDemo
{
    class Vehicle
    {
        /// <summary>
        /// Number of Vehicle
        /// </summary>
        public string VehNr { get; set; }

        /// <summary>
        /// Flag: is Vehicle in Queue? + = yes, - = no
        /// </summary>
        public string Queue { get; set; }

        /// <summary>
        /// Total Queue Time Thus Far [s]
        /// </summary>
        public string QTim { get; set; }

        /// <summary>
        /// Simulation Time [s]
        /// </summary>
        public string t { get; set; }

        /// <summary>
        /// World coordinate x (vehicle rear end at the end of the time step)
        /// </summary>
        public string RworldldX { get; set; }
    }
}
