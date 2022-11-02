using GazpromTestModels;
using System;

namespace GazpromTest.Models.GazpromTestEventArgs
{
    public class CurentInfoChangedEventArgs : EventArgs
    {
        public CurentInfoChangedEventArgs(ObjectInfo info)
        {
            Info = info;
        }

        public ObjectInfo Info { get; set; }
    }
}
