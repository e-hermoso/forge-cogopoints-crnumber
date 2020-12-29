using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Aec.ApplicationServices;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

[assembly: CommandClass(typeof(CornerRecordUpdate.Commands))]
[assembly: ExtensionApplication(null)]
namespace CornerRecordUpdate
{
    #region Commands
    public class Commands
    {
        [CommandMethod("OCPWRENAME")]
        public void ListAttributes()
        {
            var civilDoc = CivilApplication.ActiveDocument;
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;

            // read input parameters from JSON file 
            dynamic dynamicResultObject = JsonConvert.DeserializeObject(File.ReadAllText("params.json"));

            try
            {
                var acDB = doc.Database;

                using (var tr = acDB.TransactionManager.StartTransaction())
                {
                    // Cast dynamicResultObject to dictionary
                    Dictionary<string, string> docNumber = dynamicResultObject as Dictionary<string, string>;

                    LayoutManager layoutMgr = LayoutManager.Current;

                    // Capture Layouts to be Checked
                    DBDictionary layoutPages = (DBDictionary)tr.GetObject(acDB.LayoutDictionaryId, OpenMode.ForRead);

                    // Capture Cogo Points 
                    CogoPointCollection cogoPointsColl = CogoPointCollection.GetCogoPoints(doc.Database);

                    // Create a temp dictionary to associate layout name to corner record number
                    Dictionary<string, string> crTempAlias = new Dictionary<string, string>();


                    // Replace matching CogoPoint names to their Doc Number counterparts from params.json
                    foreach (ObjectId cogoPointRecord in cogoPointsColl)
                    {
                        CogoPoint cogoPointItem = tr.GetObject(cogoPointRecord, OpenMode.ForRead) as CogoPoint;

                        Match cogoMatch = Regex.Match(cogoPointItem.PointName, "^(\\s*cr\\s*\\d\\d*)$", RegexOptions.IgnoreCase);

                        if (cogoMatch.Success)
                        {
                            string cogoPointChecked = cogoPointItem.PointName.Trim().ToString().ToLower().Replace(" ", "");

                            if (dynamicResultObject.ContainsKey(cogoPointChecked.ToString()))
                            {
                                //ed.WriteMessage("\n" + dynamicResultObject[cogoPointChecked.ToString()]);
                                cogoPointItem.PointName = (dynamicResultObject[cogoPointChecked.ToString()]);
                            }
                        }
                    }

                    // Match the "Corner Record #" field from the Points table
                    foreach (ObjectId cogoPointRecord in cogoPointsColl)
                    {
                        CogoPoint cogoPointItem = tr.GetObject(cogoPointRecord, OpenMode.ForRead) as CogoPoint;

                        foreach (UDPClassification udpClassification in civilDoc.PointUDPClassifications)
                        {
                            if (udpClassification.Name.ToString() == "Corner Record No.")
                            {
                                //ed.WriteMessage("\n\nUDP Classification: {0}\n", udpClassification.Name);
                                foreach (UDP udp in udpClassification.UDPs)
                                {
                                    if (udp.Name.ToString() == "Corner Record #")
                                    {
                                        var udpType = udp.GetType().Name;
                                        if (udpType == "UDPString")
                                        {
                                            UDPString udpString = (UDPString)udp;
                                            var cornerRecordNumber = cogoPointItem.GetUDPValue(udpString);
                                            Match crNumMatch = Regex.Match(cornerRecordNumber, "^(\\s*cr\\s*\\d\\d*)$", RegexOptions.IgnoreCase);

                                            if (crNumMatch.Success)
                                            {
                                                string crNumChecked = cornerRecordNumber.Trim().ToString().ToLower().Replace(" ", "");

                                                if (dynamicResultObject.ContainsKey(crNumChecked.ToString()))
                                                {
                                                    //ed.WriteMessage("\n" + dynamicResultObject[crNumChecked.ToString()]);
                                                    cogoPointItem.SetUDPValue(udpString, dynamicResultObject[crNumChecked.ToString()].ToString());
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }


                    foreach (DBDictionaryEntry layoutPage in layoutPages)
                    {
                        var layoutUnchecked = layoutPage.Value.GetObject(OpenMode.ForRead) as Layout;
                        var isModelSpace = layoutUnchecked.ModelType;

                        ObjectIdCollection textObjCollection = new ObjectIdCollection();

                        if (isModelSpace != true)
                        {
                            Match layoutNameMatch = Regex.Match(layoutUnchecked.LayoutName, "^(\\s*cr\\s*\\d\\d*)$",
                               RegexOptions.IgnoreCase);

                            if (layoutNameMatch.Success)
                            {
                                string layoutChecked = layoutUnchecked.LayoutName.Trim().ToString().ToLower().Replace(" ", "");

                                // for each Form_Result.Value in Doc_Num_Test.json file, 
                                // if the file, if the SF_Json_Key matches layoutPage,
                                // rename the layoutPage.LayoutName with the SF_Json_Key.Key["Doc_Num"]
                                if (dynamicResultObject.ContainsKey(layoutChecked.ToString()))
                                {
                                    //ed.WriteMessage("\n" + dynamicResultObject[layoutChecked.ToString()]);

                                    crTempAlias.Add(layoutUnchecked.LayoutName.ToString(), layoutChecked);
                                }
                            }
                        }
                    }

                    // Iterate through temp alias dictionary and update the layout name 
                    foreach (KeyValuePair<string, string> entry in crTempAlias)
                    {
                        if (layoutPages.Contains(entry.Key))
                        {
                            layoutMgr.RenameLayout(entry.Key, dynamicResultObject[entry.Value].ToString());
                        }
                    }

                    tr.Commit();
                }
                acDB.SaveAs("outputFile.dwg", DwgVersion.Current);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("Exception: " + ex.Message);
            }
        }
    }
    #endregion
}
