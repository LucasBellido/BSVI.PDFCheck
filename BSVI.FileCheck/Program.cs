using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using MFilesAPI;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Xml;
using System.Xml.Linq;
using System.Reflection;

namespace BSVI.FileCheck
{
    class Program
    {
        static Aspose.Pdf.License license;
        static void Main(string[] args)
        {
            string outputPath = "";

            try
            {
                Console.WriteLine("Please specify the full path to the XML Vault Configuration File");
                var xmlPath = Console.ReadLine();

                XmlDocument xdoc = new XmlDocument();
                xdoc.Load(xmlPath);
                string response = xdoc.InnerXml;
                XDocument doc = XDocument.Parse(response);

                string ErrMsg = "";
                string VaultLogin = (string)doc.Descendants("VaultUsername").ElementAt(0);
                string VaultPassword = (string)doc.Descendants("VaultPassword").ElementAt(0);
                bool isWindowsUser = (bool)doc.Descendants("isWindowsUser").ElementAt(0);
                string VaultEndpoint = (string)doc.Descendants("VaultEndpoint").ElementAt(0);
                string VaultProtocol = (string)doc.Descendants("VaultProtocol").ElementAt(0);
                string VaultGUID = (string)doc.Descendants("VaultGUID").ElementAt(0);
                string VaultServer = (string)doc.Descendants("VaultServer").ElementAt(0);
                string ClassGroupFilter = (string)doc.Descendants("ClassGroups").ElementAt(0);
                string ClassFilter = (string)doc.Descendants("Classes").ElementAt(0);
                outputPath = (string)doc.Descendants("OutputPath").ElementAt(0);
                string restAPILink = "";
                Vault vault;

                vault = VaultConnect.ConnectServer(ref ErrMsg, VaultGUID, VaultLogin, VaultPassword, isWindowsUser, VaultServer, VaultEndpoint, VaultProtocol);
                var classgroups = vault.ClassGroupOperations.GetClassGroups(0);

                IDictionary<string, int> classGroupDictionary = new Dictionary<string, int>();

                foreach (ClassGroup classgroup in classgroups)
                {
                    if (!classGroupDictionary.ContainsKey(classgroup.Name))
                    {
                        classGroupDictionary.Add(classgroup.Name, classgroup.ID);
                    }
                }

                var classes = vault.ClassOperations.GetObjectClasses(0);

                IDictionary<string, int> classDictionary = new Dictionary<string, int>();

                foreach (ObjectClass clas in classes)
                {
                    if (!classDictionary.ContainsKey(clas.Name))
                    {
                        classDictionary.Add(clas.Name, clas.ID);
                    }
                }

                List<int> classGroupList = new List<int>();
                List<int> classList = new List<int>();

                var classGroupsStrings = ClassGroupFilter.Split(',');
                var classStrings = ClassFilter.Split(',');

                if (classStrings[0] != "" && classGroupsStrings[0] != "")
                {
                    throw new Exception("A filter for class groups and classes cannot be specified at the same time");
                }


                IDictionary<string, int> classGroupCounter = new Dictionary<string, int>();
                IDictionary<string, int> classCounter = new Dictionary<string, int>();

                if (classGroupsStrings[0] != "")
                {
                    foreach (var classGroupString in classGroupsStrings)
                    {
                        var trimmedString = classGroupString.Trim();
                        classGroupList.Add(classGroupDictionary[trimmedString]);
                        classGroupCounter.Add(trimmedString, 0);
                    }
                }
                if (classStrings[0] != "")
                {
                    foreach (var classString in classStrings)
                    {
                        var trimmedString = classString.Trim();
                        classList.Add(classDictionary[trimmedString]);
                        classCounter.Add(trimmedString, 0);
                    }
                }



                if (vault == null)
                {
                    throw new Exception("Problems connecting to Vault, please verify configuration settings");
                }

                var vaultName = vault.Name;
                CheckFiles(vault, outputPath, VaultLogin, VaultPassword, restAPILink, VaultServer, classList, classGroupList, classCounter, classGroupCounter);
            }
            catch (Exception e)
            {
                using (StreamWriter sw = System.IO.File.CreateText(outputPath))
                {
                    sw.WriteLine("Error Occured: ");
                    if (e.Message.Contains("given key"))
                    {
                        sw.WriteLine("The specified class group or class could not be found. Please verify the spelling.");

                    }
                    else
                    {
                        sw.WriteLine(e.Message);

                    }

                }
            }
        }

        public static void CheckFiles(Vault vault, string outputPath, string username, string password, string restAPILink, string VaultServer, List<int> classes, List<int> classGroups, IDictionary<string, int> classCounter, IDictionary<string, int> classGroupCounter)
        {
            try
            {
                getLicense();


                List<ObjectVersion> corruptDocuments = new List<ObjectVersion>();

                var vaultGuid = vault.GetGUID();
                var temps = findAllDocuments(vault, classGroups, classes);


                //   var authToken = getAuthToken(username, password, vaultGuid, restAPILink);

                var tempPath = Path.GetTempPath();
                var counter = 0;


                foreach (ObjectVersion temp in temps)
                {

                    var objectID = temp.OriginalObjID.ID;
                    var oObjectFiles = temp.Files;

                    counter++;

                    if (classes.Count > 0)
                    {
                        var objVer = temp.ObjVer;

                        ObjID objID = objVer.ObjID;
                        var latestObjVer = vault.ObjectOperations.GetLatestObjVer(objID, false);

                        var properties = vault.ObjectOperations.GetObjectVersionAndProperties(latestObjVer, false).Properties;
                        var classPV = properties.SearchForProperty(100).TypedValue.DisplayValue;
                        Console.WriteLine("Analyzing Document " + counter + " out of " + temps.Count + ".  " + classPV + ", " + DateTime.Now);
                        classCounter[classPV]++;

                    }
                    else if (classGroups.Count > 0)
                    {
                        var objVer = temp.ObjVer;

                        ObjID objID = objVer.ObjID;
                        var latestObjVer = vault.ObjectOperations.GetLatestObjVer(objID, false);

                        var properties = vault.ObjectOperations.GetObjectVersionAndProperties(latestObjVer, false).Properties;
                        var classGroupPV = properties.SearchForProperty(101).TypedValue.DisplayValue;

                        Console.WriteLine("Analyzing Document " + counter + " out of " + temps.Count + ".  " + classGroupPV + ", " + DateTime.Now);
                        classGroupCounter[classGroupPV]++;
                    }
                    else
                    {
                        Console.WriteLine("Analyzing Document " + counter + " out of " + temps.Count + ".  " + DateTime.Now);
                    }

                    foreach (ObjectFile oObjectFile in oObjectFiles)
                    {
                        var fileID = oObjectFile.ID;

                        var fileExtension = oObjectFile.Extension;

                        var szTargetPath = Path.Combine(tempPath, oObjectFile.GetNameForFileSystem());

                        if (fileExtension == "pdf")
                        {
                            vault.ObjectFileOperations.DownloadFile(fileID, oObjectFile.Version, szTargetPath);
                            var isValid = readPDF(szTargetPath);
                            if (!isValid)
                            {
                                corruptDocuments.Add(temp);
                            }
                            System.IO.File.Delete(szTargetPath);

                        }
                        /*  else if (fileExtension == "docx")
                         {
                             vault.ObjectFileOperations.DownloadFile(fileID, oObjectFile.Version, szTargetPath);
                             var isValid = validateWordDocs(szTargetPath);
                             if (!isValid)
                             {
                                 corruptDocuments.Add(temp);
                             }
                             File.Delete(szTargetPath);
                         }
                        else
                         {
                             var isValid = requestPreview(authToken, restAPILink, objectID.ToString(), fileID.ToString());
                             if (!isValid)
                             {
                                 corruptDocuments.Add(temp);
                             }

                         }
                         */
                    }

                }

                using (StreamWriter sw = System.IO.File.CreateText(outputPath))
                {
                    sw.WriteLine("CORRUPTED FILES REPORT");
                    sw.WriteLine("Server | " + VaultServer);
                    sw.WriteLine("Vault | " + vault.Name);
                    sw.WriteLine("Date | " + DateTime.Now);
                    sw.WriteLine("Number of Documents | " + temps.Count);
                    sw.WriteLine("Number of Corrupt Documents | " + corruptDocuments.Count);
                    sw.WriteLine(" ");
                    sw.WriteLine("Corrupt Documents: ");
                    sw.WriteLine(" ");
                    foreach (ObjectVersion corruptDocument in corruptDocuments)
                    {

                        var objVer = corruptDocument.ObjVer;

                        ObjID objID = objVer.ObjID;
                        var latestObjVer = vault.ObjectOperations.GetLatestObjVer(objID, false);

                        var properties = vault.ObjectOperations.GetObjectVersionAndProperties(latestObjVer, false).Properties;
                        var classPV = properties.SearchForProperty(100).TypedValue.DisplayValue;

                        sw.WriteLine(classPV + " | " + corruptDocument.DisplayID + " | " + corruptDocument.Title);
                    }

                    sw.WriteLine(" ");
                    foreach (var temp in classCounter)
                    {
                        sw.WriteLine(temp.Value + " document(s) of Class: " + temp.Key + " analyzed.");
                    }
                    foreach (var temp in classGroupCounter)
                    {
                        sw.WriteLine(temp.Value + " document(s) of Class Group: " + temp.Key + " analyzed.");
                    }

                }
                
            }
            catch (Exception e)
            {
                var temp = e.Message;
            }
        }
        private static void getLicense()
        {
            license = new Aspose.Pdf.License();
            var assembly = Assembly.GetExecutingAssembly();
            var stream = assembly.GetManifestResourceStream("BSVI.FileCheck.Aspose.Total.lic");
            if (stream == null)
            {
                throw new Exception("Invalid aspose license");
            }
            license.SetLicense(stream);
        }

        public static string getAuthToken(string username, string password, string vaultGuid, string RESTAPILink)
        {
            var jsonSerializer = JsonSerializer.CreateDefault();

            // Create the authentication details.
            var auth = new
            {
                Username = username,
                Password = password,
                VaultGuid = vaultGuid // Use GUID format with {braces}.
            };


            var authenticationRequest = (HttpWebRequest)WebRequest.Create(RESTAPILink + "/server/authenticationtokens.aspx");
            authenticationRequest.Method = "POST";

            // Add the authentication details to the request stream.
            using (var streamWriter = new StreamWriter(authenticationRequest.GetRequestStream()))
            {
                using (var jsonTextWriter = new JsonTextWriter(streamWriter))
                {
                    jsonSerializer.Serialize(jsonTextWriter, auth);
                }
            }

            // Execute the request.
            var authenticationResponse = (HttpWebResponse)authenticationRequest.GetResponse();

            // Extract the authentication token.
            string authenticationToken = null;
            using (var streamReader = new StreamReader(authenticationResponse.GetResponseStream()))
            {
                using (var jsonTextReader = new JsonTextReader(streamReader))
                {
                    authenticationToken = ((dynamic)jsonSerializer.Deserialize(jsonTextReader)).Value;
                }
            }

            return authenticationToken;
        }

        public static bool requestPreview(string authToken, string RESTAPILink, string objID, string fileID)
        {
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(RESTAPILink + "/objects/0/" + objID + "/latest/files/" + fileID + "/content?format=PDF");
                request.Method = "GET";

                // Ensure we set the authentication header.
                request.Headers.Add("X-Authentication", authToken);

                // Execute the request.
                var response = (HttpWebResponse)request.GetResponse();

                return true;
            }
            catch (Exception e)
            {
                return false;
            }

        }

        public static bool readPDF(string path)
        {
            try
            {
                using (Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(path))
                {
                    return true;
                }
            }

            catch (Aspose.Pdf.InvalidPasswordException e)
            {
                return true;
            }
            catch(Exception e)
            {
                return false;
            }
        }

        public static bool validateWordDocs(string filePath)
        {
            try
            {
                OpenXmlValidator validator = new OpenXmlValidator();
                using (var doc = WordprocessingDocument.Open(filePath, true))
                {
                    bool isValid = validator.Validate(doc).Count() == 0;
                    return true;
                }
            }
            catch (Exception e)
            {
                return false;
            }
        }

        public static ObjectVersions findAllDocuments(Vault vault, List<int> classGroupNumbers, List<int> classNumbers)
        {

            var classGroupArray = classGroupNumbers.ToArray();
            var classArray = classNumbers.ToArray();

            var findLatestPart = new SearchConditions();

            {
                var condition = new SearchCondition();
                condition.Expression.SetStatusValueExpression(MFStatusType.MFStatusTypeObjectTypeID);
                condition.ConditionType = MFConditionType.MFConditionTypeEqual;
                condition.TypedValue.SetValue(MFDataType.MFDatatypeLookup, 0);
                findLatestPart.Add(-1, condition);
            }
            if (classGroupNumbers.Count > 0)
            {
                var condition = new SearchCondition();
                condition.Expression.SetPropertyValueExpression(101, MFParentChildBehavior.MFParentChildBehaviorNone);
                condition.ConditionType = MFConditionType.MFConditionTypeEqual;
                condition.TypedValue.SetValue(MFDataType.MFDatatypeMultiSelectLookup, classGroupArray);
                findLatestPart.Add(-1, condition);
            }

            if (classNumbers.Count > 0)
            {
                var condition = new SearchCondition();
                condition.Expression.SetPropertyValueExpression(100, MFParentChildBehavior.MFParentChildBehaviorNone);
                condition.ConditionType = MFConditionType.MFConditionTypeEqual;
                condition.TypedValue.SetValue(MFDataType.MFDatatypeMultiSelectLookup, classArray);
                findLatestPart.Add(-1, condition);
            }
            {

                var condition = new SearchCondition();
                condition.Expression.SetStatusValueExpression(MFStatusType.MFStatusTypeDeleted);
                condition.ConditionType = MFConditionType.MFConditionTypeEqual;
                condition.TypedValue.SetValue(MFDataType.MFDatatypeBoolean, false);
                findLatestPart.Add(-1, condition);
            }

            var newPartObject = vault.ObjectSearchOperations.SearchForObjectsByConditionsEx(findLatestPart, MFSearchFlags.MFSearchFlagNone, SortResults: false, MaxResultCount: 1000000);
            var temp = newPartObject.GetAsObjectVersions();
            return temp;

        }
    }
}
