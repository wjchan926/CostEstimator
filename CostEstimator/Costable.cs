using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.Data;
using System.Runtime.InteropServices;

namespace InvAddIn
{
    class Costable
    {
        Application invApp;
        public Document inventorDoc;
        DataTable matBreakdown;
        public double cost { get; private set; }
        //public double rawCost { get; private set; } = 0;
        //public double purchasedCost { get; private set; } = 0;

        static readonly Dictionary<string, double> materialMap = new Dictionary<string, double>()
        {
            { "1008 / 1010 PLAIN STEEL", 0.06 },
            { "304 STAINLESS STEEL", 1.98 },
            { "316 STAINLESS STEEL", 2.81 },
            { "Galvanized Steel", 0.72 }
        };
        
        Costable()
        {

        }

        public Costable(Application currentApp, Document currentDoc)
        {
            invApp = currentApp;
            inventorDoc = currentDoc;
            initializeMatBreakdown();
        }

        private void initializeMatBreakdown()
        {
            matBreakdown = new DataTable();
            matBreakdown.Columns.Add("Description", typeof(string));
            matBreakdown.Columns.Add("Weight", typeof(decimal));
            matBreakdown.Columns.Add("Material", typeof(string));
            matBreakdown.Columns.Add("Cost", typeof(decimal));
        }

        public void CostOf(Document bomItem)
        {
            double mass;
            string material;
                                    
            if (bomItem.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                
                PartDocument partDoc = (PartDocument)bomItem;
                PartComponentDefinition partCompDef = partDoc.ComponentDefinition;
                PropertySets propertySets = partDoc.PropertySets;
                PropertySet propertySet = propertySets["Design Tracking Properties"];    
                Property costProp = propertySet["Cost"];

                if (partCompDef.BOMStructure == BOMStructureEnum.kPurchasedBOMStructure || partCompDef.BOMStructure == BOMStructureEnum.kInseparableBOMStructure)
                {
                    cost = Convert.ToDouble(costProp.Value);
                   // purchasedCost = Convert.ToDouble(costProp.Value);

                    Marshal.ReleaseComObject(costProp);
                    costProp = null;

                    Marshal.ReleaseComObject(propertySet);
                    propertySet = null;

                    Marshal.ReleaseComObject(propertySets);
                    propertySets = null;

                    Marshal.ReleaseComObject(partCompDef);
                    partCompDef = null;

                    Marshal.ReleaseComObject(partDoc);
                    partDoc = null;

                    return;
                }

                Property materialProp = propertySet["Material"];
                
                // Mass of Part
                mass = partCompDef.MassProperties.Mass;

                // Material of Part
                material = materialProp.Value.ToString();

                // Cost
                try
                {
                    cost = mass * materialMap[material];
                }
                catch (Exception)
                {
                    cost = 0.0;
                }
                         
                costProp.Value = cost;  // Set Estimated Cost

                Marshal.ReleaseComObject(materialProp);
                materialProp = null;

                Marshal.ReleaseComObject(costProp);
                costProp = null;

                Marshal.ReleaseComObject(propertySet);
                propertySet = null;

                Marshal.ReleaseComObject(propertySets);
                propertySets = null;

                Marshal.ReleaseComObject(partCompDef);
                partCompDef = null;

                Marshal.ReleaseComObject(partDoc);
                partDoc = null;

                return;
            }            

            foreach(Document nestedAsm in bomItem.ReferencedDocuments)
            {
                CostOf(nestedAsm);
            }

            if (bomItem.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                cost = 0;

                AssemblyDocument assemblyDoc = (AssemblyDocument)bomItem;
                PropertySets assemblySets = assemblyDoc.PropertySets;
                PropertySet assemblySet = assemblySets["Design Tracking Properties"];
                Property asmCost = assemblySet["Cost"];

                AssemblyComponentDefinition asmCompDef = assemblyDoc.ComponentDefinition;
                ComponentOccurrences occurances = asmCompDef.Occurrences;

                foreach (Document nestedPart in bomItem.ReferencedDocuments)
                {
                    try
                    {
                        PropertySets propertySets = nestedPart.PropertySets;
                        PropertySet propertySet = propertySets["Design Tracking Properties"];
                        Property costProp = propertySet["Cost"];

                        cost = cost + Convert.ToDouble(costProp.Value.ToString()) * GetOccurances(nestedPart, occurances);

                        Marshal.ReleaseComObject(costProp);
                        costProp = null;

                        Marshal.ReleaseComObject(propertySet);
                        propertySet = null;

                        Marshal.ReleaseComObject(propertySets);
                        propertySets = null;

                    }
                    catch
                    {
                        continue;
                    }
  
                }

                asmCost.Value = cost;

                Marshal.ReleaseComObject(occurances);
                occurances = null;

                Marshal.ReleaseComObject(asmCompDef);
                asmCompDef = null;

                Marshal.ReleaseComObject(asmCost);
                asmCost = null;

                Marshal.ReleaseComObject(assemblySet);
                assemblySet = null;

                Marshal.ReleaseComObject(assemblySets);
                assemblySets = null;

                Marshal.ReleaseComObject(assemblyDoc);
                assemblyDoc = null;
            }
        }

        private int GetOccurances(Document doc, ComponentOccurrences occurances)
        {
            int numOccurances = 0;

            foreach(ComponentOccurrence occurance in occurances)
            {
                string partName = (occurance.Name).Substring(0, (occurance.Name).LastIndexOf(':'));
                if ( partName == doc.DisplayName)
                {
                    numOccurances++;
                }
            }

            return numOccurances;
        }
    }

}
