using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GazpromTestModels
{
    public class ObjectInfo
    {
        public ObjectInfo()
        {
        }

        public ObjectInfo(string name, double distance, double angle, double width, double heigth, bool isDefect)
        {
            Name = name;
            Distance = distance;
            Angle = angle;
            Width = width;
            Heigth = heigth;
            IsDefect = isDefect;
        }

        public string Name { get; set; }
        public double Distance { get; set; }
        public double Angle { get; set; }
        public double Width { get; set; }
        public double Heigth { get; set; }
        public bool IsDefect { get; set; }

    }
}
