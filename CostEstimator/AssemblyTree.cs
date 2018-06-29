using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;


namespace InvAddIn
{
    class AssemblyTree
    {
        class PartNode : INode
        {
            public PartDocument data;
            public List<INode> children = null;

            public PartNode(Document doc)
            {
                SetData(doc);
            }

            public PartNode(PartDocument doc)
            {
                data = doc;
            }

            public void SetData(Document doc)
            {
                data = (PartDocument)doc;
            }

            public List<INode> GetChildren()
            {
                return children;
            }
        }

        class AsmNode : INode
        {
            public AssemblyDocument data;
            public List<INode> children;

            public AsmNode(Document doc)
            {
                SetData(doc);
                children = new List<INode>();
            }

            public AsmNode(AssemblyDocument doc)
            {
                data = doc;
            }

            public void SetData(Document doc)
            {
                data = (AssemblyDocument)doc;
            }

            public List<INode> GetChildren()
            {
                return children;
            }
        }

        // Master Assembly
        INode root;

        public AssemblyTree(Document doc)
        {
            // Create Tree
            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                root = new AsmNode(doc);
            }
            else
            {
                root = new PartNode(doc);
            }       
                                                     
        }

        public AssemblyTree()
        {
            root = null;
        }
        
        public void push(Document doc)
        {
            if()
        }
        

    }
}
