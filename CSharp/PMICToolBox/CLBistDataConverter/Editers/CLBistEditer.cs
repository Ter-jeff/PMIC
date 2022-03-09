using CLBistDataConverter.DataStructures;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CLBistDataConverter
{
    public class CLBistEditer
    {
        public void Edit(List<CLBistDie> clBistDieDatalst)
        {
            clBistDieDatalst.ForEach(s => s.CLBistSitelst.ForEach(m => m.EditSiteClbistOutputRows()));
        }
    }
}
