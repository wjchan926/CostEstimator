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
        public DataTable matBreakdown { get; private set; }
        public double cost { get; private set; }
        string matFilepath = @"\\MSW-FP1\Shared\Marlin Software Output\Cost Estimator\MaterialPrices.csv";
        //public double rawCost { get; private set; } = 0;
        //public double purchasedCost { get; private set; } = 0;
               
        static Dictionary<string, double> materialMap = new Dictionary<string, double>();
        
        Costable()
        {

        }

        public Costable(Application currentApp, Document currentDoc)
        {
            CreateMaterialDB();
            invApp = currentApp;
            inventorDoc = currentDoc;
            initializeMatBreakdown();
        }

        private void CreateMaterialDB()
        {
            materialMap = System.IO.File.ReadLines(matFilepath).Select(line => line.Split(',')).ToDictionary(line => line[0], line => Convert.ToDouble(line[1]));
        }

        private void initializeMatBreakdown()
        {
            matBreakdown = new DataTable();
            matBreakdown.Columns.Add("Qty", typeof(int));
            matBreakdown.Columns.Add("Description", typeof(string));
      //      matBreakdown.Columns.Add("Weight", typeof(decimal));
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
                if (materialProp.Value.ToString().Contains("MESH"))
                {
                    material = "MESH";
                }
                else if(materialProp.Value.ToString().Contains("EXPANDED METAL"))
                {
                    material = "EXPANDED METAL";
                }
                else
                {
                    material = materialProp.Value.ToString();
                }


                // Cost
                try
                {
                    // Multiple by 2.20462262 for kg to lbs conversion
                    cost = mass * 2.20462262 * materialMap[material];
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

        public void UpdateMaterial(Document bomItem)
        {
            AssemblyDocument asmDoc;

            if (bomItem.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                asmDoc = (AssemblyDocument)bomItem;

                AssemblyComponentDefinition compDef = asmDoc.ComponentDefinition;
                BOM bom = compDef.BOM;
                bom.StructuredViewFirstLevelOnly = true;

                bom.StructuredViewEnabled = true;

                BOMView bomView = bom.BOMViews["Structured"];

                QueryBOM(bomView.BOMRows);
            }
        }
        
        private void QueryBOM(BOMRowsEnumerator bomRows)
        {
            for (int i = 1; i < bomRows.Count + 1; i++)
            {
                try
                {
                    BOMRow bomRow = bomRows[i];
                               
                    ComponentDefinition compDef = bomRow.ComponentDefinitions[1];

                    Document doc = (Document)compDef.Document;
                    PropertySets propertySets = doc.PropertySets;
                    PropertySet propertySet = propertySets["Design Tracking Properties"];

                    Property descProp = propertySet["Description"];
                    //Property weightProp
                    Property materialProp = propertySet["Material"];
                    Property costProp = propertySet["Cost"];

                    matBreakdown.Rows.Add(bomRow.ItemQuantity, descProp.Value, materialProp.Value, costProp.Value);

                    Marshal.ReleaseComObject(costProp);
                    costProp = null;
                    Marshal.ReleaseComObject(materialProp);
                    materialProp = null;
                    Marshal.ReleaseComObject(descProp);
                    descProp = null;
                    Marshal.ReleaseComObject(propertySet);
                    propertySet = null;
                    Marshal.ReleaseComObject(propertySets);
                    propertySets = null;
                    Marshal.ReleaseComObject(doc);
                    doc = null;
                    Marshal.ReleaseComObject(compDef);
                    compDef = null;
                    Marshal.ReleaseComObject(bomRow);
                    bomRow = null;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }
        }        

        private int GetOccurances(Document doc, ComponentOccurrences occurances)
        {
            int numOccurances = 0;

            foreach(ComponentOccurrence occurance in occurances)
            {
                string partName = occurance.Name;

                if (occurance.Name.Contains(':'))
                {
                    partName = (occurance.Name).Substring(0, (occurance.Name).LastIndexOf(':'));
                }
                else if (occurance.Name.Contains('_'))
                {
                    partName = (occurance.Name).Substring(0, (occurance.Name).LastIndexOf('_'));
                }
                else if (occurance.Name.Contains('-'))
                {
                    partName = (occurance.Name).Substring(0, (occurance.Name).LastIndexOf('-'));
                }

                if (partName == doc.DisplayName)
                {
                    numOccurances++;
                }

            }

            return numOccurances;
        }
    }

}
