﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvAddIn
{
    interface INode
    {
        void SetData(Inventor.Document doc);
        List<INode> GetChildren(); 
    }


}
