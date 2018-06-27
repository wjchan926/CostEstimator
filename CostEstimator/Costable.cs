using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace InvAddIn
{
    class Costable
    {
        Application invApp;
        Document inventorDoc;
        BOM bom;

        Costable()
        {

        }

        Costable(Application currentApp, dynamic currentDoc)
        {
            invApp = currentApp;
            inventorDoc = currentDoc;
            bom = currentDoc.Compoonent
        }

    }
}
