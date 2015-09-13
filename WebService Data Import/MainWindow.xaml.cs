using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.Collections.ObjectModel;

using System.Web;


using System.Net;

using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using System.IO;

using System.Collections;


using System.CodeDom.Compiler;
using System.Globalization;
using System.Reflection;
using System.ServiceModel;
using System.ServiceModel.Description;

using Microsoft.Win32;

namespace WebService_Data_Import
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private int iDataVerticalStartPosition = 5;
        private int iDataHorizontalStartPosition = 5;

        private int giMaxColumnPosition = 16384;
        private int giMaxRowPosition = 1048576;

        private string gstrResultColumn = "A";
        private string gstrErrorCountCell = "B2";
        private string gstrStartTimeCell = "B3";
        private string gstrEndTimeCell = "D3";
        private string gstrDurationCell = "F3";

        public MainWindow()
        {
            InitializeComponent();

            this.Title += " " + Assembly.GetExecutingAssembly().GetName().Version;
        }

        static string GetColumnLetter(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];

            return value;
        }

        private string retrieveValue(Microsoft.Office.Interop.Excel.Range arngCell)
        {
            string result = null;

            if (null != arngCell.Value2)
            {
                string strValue = arngCell.Value2.ToString();
            }

            return (result);
        }

        private void btnGenerate_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofdOpenFileDialog = new OpenFileDialog();

            ofdOpenFileDialog.Filter = "Excel Files (*.xlsx, *.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";
            ofdOpenFileDialog.Multiselect = false;
            if(true == ofdOpenFileDialog.ShowDialog())
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

                if (null != excel)
                {
                    excel.Visible = true;

                    Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Open(ofdOpenFileDialog.FileName);
                    if (null != workbook)
                    {
                        Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Sheets["Parameters"];
                        if (null != worksheet)
                        {
                            Microsoft.Office.Interop.Excel.Range rangeWebService = worksheet.get_Range("C3");
                            Microsoft.Office.Interop.Excel.Range rangeUsername = worksheet.get_Range("C4");
                            Microsoft.Office.Interop.Excel.Range rangePassword = worksheet.get_Range("C5");
                            Microsoft.Office.Interop.Excel.Range rangeMethod = worksheet.get_Range("C7");

                            Microsoft.Office.Interop.Excel.Range rangeDataStartRow = worksheet.get_Range("C11");
                            Microsoft.Office.Interop.Excel.Range rangeDataStartColumn = worksheet.get_Range("C10");

                            Microsoft.Office.Interop.Excel.Range rangeResultColumn = worksheet.get_Range("C16");
                            Microsoft.Office.Interop.Excel.Range rangeErrorCountCell = worksheet.get_Range("C17");
                            Microsoft.Office.Interop.Excel.Range rangeStartTimeCell = worksheet.get_Range("C18");
                            Microsoft.Office.Interop.Excel.Range rangeEndTimeCell = worksheet.get_Range("C19");
                            Microsoft.Office.Interop.Excel.Range rangeDurationCell = worksheet.get_Range("C20");


                            string strStartRow = rangeDataStartRow.Value2.ToString();
                            string strStartColumn = rangeDataStartColumn.Value2.ToString();

                            int.TryParse(strStartRow, out iDataVerticalStartPosition);
                            int.TryParse(strStartColumn, out iDataHorizontalStartPosition);

                            string strURI = rangeWebService.Value;
                            string strUsername = rangeUsername.Value;
                            string strPassword = rangePassword.Value;
                            string strMethod = rangeMethod.Value;


                            generateSpreadsheet(strURI, strUsername, strPassword, strMethod, workbook);
                        }
                        else
                        {
                            Console.WriteLine("Failed to retrieve Parameters sheet");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Failed to open " + ofdOpenFileDialog.FileName);
                    }
                    excel.Quit();
                }
            }

        }

        /// <summary>
        /// this will return a list of types, in the order that they need to
        /// be created (there are subtypes)
        /// </summary>
        /// <param name="atpAllTypes"></param>
        /// <param name="astrPropertyTypeName"></param>
        /// <returns></returns>
        List<Type> getPropertyRecursive(Type[] atpAllTypes, string astrPropertyTypeName)
        {
            List<Type> result = null, lsWorking = new List<Type>();

            var paramType = atpAllTypes.Where(t => t.Name == astrPropertyTypeName);
            if (Enumerable.Count(paramType) > 0)
            {
                Type tpParameterType = paramType.First();
                lsWorking.Add(tpParameterType);
                PropertyInfo[] piProperties = tpParameterType.GetProperties();
                if(null != piProperties && piProperties.Count() > 0)
                {
                    // currentProperty.PropertyType
                    //if(piProperties.Count() == 1 && false == isPrimative(piProperties.First().PropertyType)) // piProperties.First().PropertyType.IsPrimitive == false && piProperties.First().PropertyType != typeof(string))
                    //{
                    //    // we need to be recursive - we should really be smarter
                    //    List<Type> lsNewTypes = getPropertyRecursive(atpAllTypes, piProperties.First().PropertyType.Name);
                    //    if(null != lsNewTypes && lsNewTypes.Count > 0)
                    //    {
                    //        lsWorking.AddRange(lsNewTypes);
                    //    }
                    //}
                    foreach(PropertyInfo currentProperty in piProperties)
                    {
                        if(false == isPrimative(currentProperty.PropertyType))
                        {
                            List<Type> lsNewTypes = getPropertyRecursive(atpAllTypes, currentProperty.PropertyType.Name);
                            if (null != lsNewTypes && lsNewTypes.Count > 0)
                            {
                                lsWorking.AddRange(lsNewTypes);
                            }
                        }
                    }
                }
            }

            if(lsWorking.Count > 0)
            {
                result = lsWorking;
            }

            return (result);
        }

        /// <summary>
        /// this will retrieve the class Type for the parameter of the method we want to call
        /// </summary>
        /// <param name="atpTypes">an array of all of the types in the compiled assembly</param>
        /// <param name="astrMethod">name of the method</param>
        /// <returns></returns>
        List<Type> getArgumentClassType(Type[] atpTypes, string astrMethod)
        {
            List<Type> result = null;

            foreach (var currentMethod in atpTypes)
            {
                if (null != currentMethod && null != currentMethod.BaseType)
                {
                    MethodInfo miMethod = currentMethod.GetMethod(astrMethod);
                    if (null != miMethod)
                    {
                        ParameterInfo[] parameters = miMethod.GetParameters();
                        if (null != parameters)
                        {
                            foreach (ParameterInfo currentMethodArgumentParameter in parameters)
                            {
                                if (null != currentMethodArgumentParameter)
                                {
                                    if (currentMethodArgumentParameter.Name != "lws" && currentMethodArgumentParameter.Name != "mws")
                                    {
                                        Console.WriteLine("Method: " + currentMethod.Name + " Type: " + currentMethodArgumentParameter.ParameterType + " " + currentMethodArgumentParameter.Name);

                                        result = getPropertyRecursive(atpTypes, currentMethodArgumentParameter.ParameterType.ToString());
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return (result);
        }

        private List<Parameter> getParameterColumns(Microsoft.Office.Interop.Excel.Worksheet awsData, int aiParameterTitleStartX, int aiParameterTitleStartY)
        {
            List<Parameter> result = null, lsParameters = new System.Collections.Generic.List<Parameter>();

            int iXPos = aiParameterTitleStartX;
            int iYPos = aiParameterTitleStartY;


            for (int i = iXPos; i < giMaxColumnPosition; i++)
            {
                string strColumnLetter = GetColumnLetter(i);

                string strName = awsData.get_Range(strColumnLetter + (iYPos-2)).Value2;

                if(false == String.IsNullOrWhiteSpace(strName))
                {
                    lsParameters.Add(new Parameter() { Name = strName, Value = strColumnLetter });
                }
                else
                {
                    break;
                }
            }

            if (lsParameters.Count > 0)
            {
                result = lsParameters;
            }


            return (result);
        }

        private object createObjectHierachy(List<Type> alsTypeHeirachy, List<object> alsObjectsCreated, int aiDepth)
        {
            object result = null;

            Type tpCurrentObjectType = alsTypeHeirachy[aiDepth];

            result = System.Activator.CreateInstance(tpCurrentObjectType);

            alsObjectsCreated.Add(result);

            if (aiDepth < alsTypeHeirachy.Count - 1)
            {
                object pohHierachy = createObjectHierachy(alsTypeHeirachy, alsObjectsCreated, aiDepth + 1);

                if (null != pohHierachy)
                {
                    PropertyInfo[] piProperties = tpCurrentObjectType.GetProperties();
                    if(null != piProperties)
                    {
                        PropertyInfo piProperty = piProperties.Where(p => p.PropertyType == pohHierachy.GetType()).First();
                        tpCurrentObjectType.InvokeMember(piProperty.Name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, result, new object[] { pohHierachy });
                    }
                }
            }

            return (result);
        }

        private void callWebService(compiledAssembly compiledAsm, string astrMethod, List<Parameter> alsParameters, Microsoft.Office.Interop.Excel.Worksheet awsData, int aiStartRow)
        {
            CompilerResults compilerResult = compiledAsm.compilerResults as CompilerResults;

            int iErrorCount = 0;
            DateTime dtStart = DateTime.Now;
            DateTime dtEnd;

            if (null != compilerResult && false == string.IsNullOrEmpty(astrMethod))
            {
                var allTypes = compilerResult.CompiledAssembly.GetTypes();

                MethodInfo methodInfo = null; //compilerResult.GetType().GetMethod(astrMethod);

                foreach (var currentMethod in allTypes)
                {
                    if (null != currentMethod && null != currentMethod.BaseType)
                    {
                        if(null != (methodInfo = currentMethod.GetMethod(astrMethod)))
                        {
                            break;
                        }
                    }
                }

                var headerInfo = allTypes.Where(t => t.Name == "lws" || t.Name == "mws").First();

                object objHeader = System.Activator.CreateInstance(headerInfo);

                // this is the type of the final type in our heirachy where we plug in the spreadsheet values
                Type tpTypeToPopulate = null;

                // argumentType = getArgumentClassType(allTypes, astrMethod);

                List<Type> lsTypeHierarchy = getArgumentClassType(allTypes, astrMethod);

                List<object> lsCallObjects = new List<object>();
                // create the object hierachy, the first object in the list should be the one
                // supplied to the webservice call.  The last should be the one that we populate
                // with data from the spreadsheet
                createObjectHierachy(lsTypeHierarchy, lsCallObjects, 0);


                if(lsCallObjects.Count > 0)
                {
                    tpTypeToPopulate = lsTypeHierarchy.Last();
                    // object objObjectToSet = lsCallObjects.Last();
                    object objObjectToPassToCall = lsCallObjects.First();

                    foreach(object objObjectToSet in lsCallObjects)
                    {
                        tpTypeToPopulate = objObjectToSet.GetType();

                        for (int i = aiStartRow; i < giMaxRowPosition; i++)
                        {
                            bool bBlank = true;

                            foreach (Parameter currentParameter in alsParameters)
                            {
                                if (null != currentParameter)
                                {
                                    var spreadSheetValue = awsData.get_Range(currentParameter.Value + i).Value2;
                                    if (null != spreadSheetValue)
                                    {
                                        // retrieve the value from the spreadsheet
                                        string strValue = awsData.get_Range(currentParameter.Value + i).Value2.ToString();

                                        PropertyInfo currentProperty = tpTypeToPopulate.GetProperty(currentParameter.Name);

                                        if (null != currentProperty)
                                        {
                                            if (false == string.IsNullOrWhiteSpace(strValue))
                                            {
                                                bBlank = false;
                                            }

                                            object objValue = strValue;

                                            if (typeof(int) == currentProperty.PropertyType)
                                            {
                                                int iValue = 0;
                                                int.TryParse(strValue, out iValue);

                                                objValue = iValue;
                                            }
                                            else if (typeof(decimal) == currentProperty.PropertyType)
                                            {
                                                decimal decValue = 0;
                                                decimal.TryParse(strValue, out decValue);

                                                objValue = decValue;
                                            }
                                            else if (typeof(double) == currentProperty.PropertyType)
                                            {
                                                double dblValue = 0;
                                                double.TryParse(strValue, out dblValue);

                                                objValue = dblValue;
                                            }
                                            else if (typeof(DateTime) == currentProperty.PropertyType)
                                            {
                                                DateTime dtValue = DateTime.MinValue;

                                                DateTime.TryParse(strValue, out dtValue);

                                                objValue = dtValue;
                                            }

                                            // save the parameter
                                            tpTypeToPopulate.InvokeMember(currentParameter.Name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, objObjectToSet, new object[] { objValue });
                                        }
                                        else
                                        {
                                            // property doesn't exist
                                        }

                                    }
                                }
                            }
                            // if we have a blank line we should stop
                            if (true == bBlank)
                            {
                                break;
                            }
                            else
                            {
                                // call to the webservice here
                                try
                                {
                                    methodInfo.Invoke(compiledAsm.instantiatedObject, new object[] { objHeader, objObjectToPassToCall });
                                    awsData.get_Range(gstrResultColumn + i).Value2 = "OK";
                                }
                                catch (Exception ex)
                                {
                                    awsData.get_Range(gstrResultColumn + i).Value2 = ex;
                                    iErrorCount++;
                                }
                            }
                        }
                    }



                    dtEnd = DateTime.Now;

                    awsData.get_Range(gstrErrorCountCell).Value2 = iErrorCount;
                    awsData.get_Range(gstrStartTimeCell).Value2 = dtStart.ToString("yyyy-MM-dd HH:mm:ss");
                    awsData.get_Range(gstrEndTimeCell).Value2 = dtEnd.ToString("yyyy-MM-dd HH:mm:ss");
                    awsData.get_Range(gstrDurationCell).Value2 = (dtEnd - dtStart).TotalSeconds;
                }


            }
        }

        /// <summary>
        /// compile our wsdl file
        /// </summary>
        /// <param name="astrWebServiceURI"></param>
        /// <param name="astrUsername"></param>
        /// <param name="astrPassword"></param>
        /// <param name="astrMethod"></param>
        /// <returns></returns>
        private compiledAssembly generateCompiledAssembly(string astrWebServiceURI, string astrUsername, string astrPassword, string astrMethod)
        {
            CompilerResults result = null;

            // Define the WSDL Get address, contract name and parameters, with this we can extract WSDL details any time
            //Uri address = new Uri("http://ifbenp.indfish.co.nz:16201/mws-ws/services/Vessel?wsdl"); //("http://localhost:64508/Service1.svc?wsdl");
            Uri address = new Uri(astrWebServiceURI);
            // For HttpGet endpoints use a Service WSDL address a mexMode of .HttpGet and for MEX endpoints use a MEX address and a mexMode of .MetadataExchange
            MetadataExchangeClientMode mexMode = MetadataExchangeClientMode.HttpGet;

            // Get the metadata file from the service.
            MetadataExchangeClient metadataExchangeClient = new MetadataExchangeClient(address, mexMode);
            metadataExchangeClient.ResolveMetadataReferences = true;

            //One can also provide credentials if service needs that by the help following two lines.
            ICredentials networkCredential = new NetworkCredential(astrUsername, astrPassword, "");
            metadataExchangeClient.HttpCredentials = networkCredential;

            //Gets the meta data information of the service.
            MetadataSet metadataSet = metadataExchangeClient.GetMetadata();

            // Import all contracts and endpoints.
            WsdlImporter wsdlImporter = new WsdlImporter(metadataSet);

            //Import all contracts.
            Collection<ContractDescription> contracts = wsdlImporter.ImportAllContracts();

            //Import all end points.
            ServiceEndpointCollection allEndpoints = wsdlImporter.ImportAllEndpoints();

            // Generate type information for each contract.
            ServiceContractGenerator serviceContractGenerator = new ServiceContractGenerator();

            //Dictinary has been defined to keep all the contract endpoints present, contract name is key of the dictionary item.
            var endpointsForContracts = new Dictionary<string, IEnumerable<ServiceEndpoint>>();

            string contractName = null;

            foreach (ContractDescription contract in contracts)
            {
                serviceContractGenerator.GenerateServiceContractType(contract);
                // Keep a list of each contract's endpoints.
                endpointsForContracts[contract.Name] = allEndpoints.Where(ep => ep.Contract.Name == contract.Name).ToList();
                contractName = contract.Name;
            }

            // Generate a code file for the contracts.
            CodeGeneratorOptions codeGeneratorOptions = new CodeGeneratorOptions();
            codeGeneratorOptions.BracingStyle = "C";

            // Create Compiler instance of a specified language.
            CodeDomProvider codeDomProvider = CodeDomProvider.CreateProvider("C#");

            // Adding WCF-related assemblies references as copiler parameters, so as to do the compilation of particular service contract.
            CompilerParameters compilerParameters = new CompilerParameters(new string[] { "System.dll", "System.ServiceModel.dll", "System.Runtime.Serialization.dll" });
            compilerParameters.GenerateInMemory = true;

            //Gets the compiled assembly.
            result = codeDomProvider.CompileAssemblyFromDom(compilerParameters, serviceContractGenerator.TargetCompileUnit);

            object proxyInstance = null;

            if (result.Errors.Count <= 0)
            {

                // Find the proxy type that was generated for the specified contract (identified by a class that implements the contract and ICommunicationbject - this is contract
                //implemented by all the communication oriented objects).
                Type proxyType = result.CompiledAssembly.GetTypes().First(t => t.IsClass && t.GetInterface(contractName) != null &&
                    t.GetInterface(typeof(ICommunicationObject).Name) != null);


                // Now we get the first service endpoint for the particular contract.
                ServiceEndpoint serviceEndpoint = endpointsForContracts[contractName].First();

                BasicHttpBinding newBinding = new BasicHttpBinding();

                newBinding.Security.Mode = BasicHttpSecurityMode.TransportCredentialOnly;
                newBinding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Basic;
                newBinding.Security.Transport.ProxyCredentialType = HttpProxyCredentialType.None;
                newBinding.Security.Message.AlgorithmSuite = System.ServiceModel.Security.SecurityAlgorithmSuite.Default;

                newBinding.Security.Message.ClientCredentialType = BasicHttpMessageCredentialType.UserName;

                serviceEndpoint.Binding = newBinding;


                // Create an instance of the proxy by passing the endpoint binding and address as parameters.
                proxyInstance = result.CompiledAssembly.CreateInstance(proxyType.Name, false, System.Reflection.BindingFlags.CreateInstance, null,
                    new object[] { serviceEndpoint.Binding, serviceEndpoint.Address }, CultureInfo.CurrentCulture, null);

                PropertyInfo clientCredentials = proxyType.GetProperty("ClientCredentials");
                if(null != clientCredentials)
                {
                    ClientCredentials credentials = new ClientCredentials();
                    credentials.UserName.UserName = astrUsername;
                    credentials.UserName.Password = astrPassword;

                    object objCredentials = clientCredentials.GetValue(proxyInstance);
                    credentials = objCredentials as ClientCredentials;

                    if (null != credentials)
                    {
                        credentials.UserName.Password = astrPassword;
                        credentials.UserName.UserName = astrUsername;
                    }
                }
            }

            compiledAssembly finalResults = new compiledAssembly() { compilerResults = result, instantiatedObject = proxyInstance };

            return (finalResults);
        }

        /// <summary>
        /// retrieve the name of the service
        /// </summary>
        /// <param name="astrURI"></param>
        /// <returns></returns>
        private string extractServiceName(string astrURI)
        {
            string result = null;

            if(false == string.IsNullOrEmpty(astrURI))
            {
                int iStartPos = astrURI.LastIndexOf("/");
                int iEndPos = astrURI.LastIndexOf("?");

                result = astrURI.Substring(iStartPos + 1, iEndPos - iStartPos - 1);
            }

            return (result);
        }

        private string extractHumanReadableType(string astrType)
        {
            string result = null;

            if (false == string.IsNullOrEmpty(astrType))
            {
                int iStartPos = astrType.LastIndexOf(".");

                result = astrType.Substring(iStartPos + 1);
            }

            return (result);
        }

        private void spreadsheetDataSetTemplate(Microsoft.Office.Interop.Excel.Worksheet worksheet, string astrMethodName)
        {
            if(null != worksheet)
            {
                Microsoft.Office.Interop.Excel.Range rangeMethod = worksheet.get_Range("A1");
                Microsoft.Office.Interop.Excel.Range rangeMethodName = worksheet.get_Range("B1");
                Microsoft.Office.Interop.Excel.Range rangeErrors = worksheet.get_Range("A2");
                Microsoft.Office.Interop.Excel.Range rangeStart = worksheet.get_Range("A3");
                Microsoft.Office.Interop.Excel.Range rangeEnd = worksheet.get_Range("C3");
                Microsoft.Office.Interop.Excel.Range rangeDuraction = worksheet.get_Range("E3");

                rangeMethod.Value2 = "Method";
                rangeMethod.Font.Bold = true;
                rangeMethodName.Value2 = astrMethodName;

                rangeDuraction.Value2 = "Duration";
                rangeDuraction.Font.Bold = true;

                rangeStart.Value2 = "Start";
                rangeStart.Font.Bold = true;

                rangeEnd.Value2 = "End";
                rangeEnd.Font.Bold = true;

                rangeErrors.Value2 = "Errors";
                rangeErrors.Font.Bold = true;
            }
        }

        /// <summary>
        /// determine if we are a primative type (really we just want to know if we are a class we need to analysis further)
        /// </summary>
        /// <param name="atpCurrentType"></param>
        /// <returns></returns>
        private bool isPrimative(Type atpCurrentType)
        {
            bool result = atpCurrentType.IsPrimitive;

            if(false == result)
            {
                if(atpCurrentType == typeof(string) || atpCurrentType == typeof(decimal))
                {
                    result = true;
                }
            }

            return (result);
        }

        private int generateSpreadsheetColumnsRecursive(PropertyInfo[] apiProperties, Microsoft.Office.Interop.Excel.Worksheet awsWorksheet, int aiCurrentColumn)
        {
            int result = aiCurrentColumn;
            foreach (PropertyInfo currentProperty in apiProperties)
            {
                if (true == isPrimative(currentProperty.PropertyType))
                {
                    Microsoft.Office.Interop.Excel.Range rangeName = awsWorksheet.get_Range(GetColumnLetter(aiCurrentColumn) + (iDataVerticalStartPosition - 2));
                    Microsoft.Office.Interop.Excel.Range rangeType = awsWorksheet.get_Range(GetColumnLetter(aiCurrentColumn) + (iDataVerticalStartPosition + 1 - 2));

                    rangeName.Value2 = currentProperty.Name;
                    rangeName.Font.Bold = true;

                    rangeType.Value2 = extractHumanReadableType(currentProperty.PropertyType.ToString());
                    rangeType.Font.Italic = true;

                    aiCurrentColumn++;
                }
                else
                {
                    //Microsoft.Office.Interop.Excel.Range rangeName = awsWorksheet.get_Range(GetColumnLetter(aiCurrentColumn) + (iDataVerticalStartPosition - 3));
                    //Microsoft.Office.Interop.Excel.Range rangeType = awsWorksheet.get_Range(GetColumnLetter(aiCurrentColumn) + (iDataVerticalStartPosition - 4));

                    //rangeName.Value2 = currentProperty.Name;
                    //rangeName.Font.Bold = true;

                    //rangeType.Value2 = extractHumanReadableType(currentProperty.PropertyType.ToString());
                    //rangeType.Font.Italic = true;
                    aiCurrentColumn = generateSpreadsheetColumnsRecursive(currentProperty.PropertyType.GetProperties(), awsWorksheet, aiCurrentColumn);

                }
            }
            result = aiCurrentColumn;
            return (result);
        }

        private void generateSpreadsheet(string astrWebServiceURI, string astrUsername, string astrPassword, string astrMethod, Microsoft.Office.Interop.Excel.Workbook workbook)
        {
            CompilerResults compilerResult = null;
            compiledAssembly compiledAsm = generateCompiledAssembly(astrWebServiceURI, astrUsername, astrPassword, astrMethod);

            if (null != (compilerResult = compiledAsm.compilerResults))
            {

                if (compilerResult.Errors.Count <= 0)
                {
                    var allTypes = compilerResult.CompiledAssembly.GetTypes();
                    List<Type> lsTypeHierarchy = getArgumentClassType(allTypes, astrMethod);

                    Type tpParametersToPopulate = lsTypeHierarchy.Last();

                    PropertyInfo[] piProperties = tpParametersToPopulate.GetProperties();

                    if(null != piProperties)
                    {

                        int i = iDataHorizontalStartPosition;
                        Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Sheets["Data"];

                        if (null != worksheet)
                        {
                            spreadsheetDataSetTemplate(worksheet, astrMethod);


                            generateSpreadsheetColumnsRecursive(piProperties, worksheet, i);
                            //foreach (PropertyInfo currentProperty in piProperties)
                            //{
                            //    Microsoft.Office.Interop.Excel.Range rangeName = worksheet.get_Range(GetColumnLetter(i) + (iDataVerticalStartPosition - 2));
                            //    Microsoft.Office.Interop.Excel.Range rangeType = worksheet.get_Range(GetColumnLetter(i) + (iDataVerticalStartPosition + 1 - 2));

                            //    rangeName.Value2 = currentProperty.Name;
                            //    rangeName.Font.Bold = true;

                            //    rangeType.Value2 = extractHumanReadableType(currentProperty.PropertyType.ToString());
                            //    rangeType.Font.Italic = true;

                            //    i++;
                            //}
                        }
                    }

                }

            }
        }


        private void btnLoad_Copy_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofdOpenFileDialog = new OpenFileDialog();

            ofdOpenFileDialog.Filter = "Excel Files (*.xlsx, *.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";
            ofdOpenFileDialog.Multiselect = false;
            if (true == ofdOpenFileDialog.ShowDialog())
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

                if (null != excel)
                {
                    excel.Visible = true;

                    Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Open(ofdOpenFileDialog.FileName);
                    if (null != workbook)
                    {
                        Microsoft.Office.Interop.Excel.Worksheet wsParameters = workbook.Sheets["Parameters"];
                        Microsoft.Office.Interop.Excel.Worksheet wsData = workbook.Sheets["Data"];

                        if (null != wsParameters && null != wsData)
                        {
                            Microsoft.Office.Interop.Excel.Range rangeWebService = wsParameters.get_Range("C3");
                            Microsoft.Office.Interop.Excel.Range rangeUsername = wsParameters.get_Range("C4");
                            Microsoft.Office.Interop.Excel.Range rangePassword = wsParameters.get_Range("C5");
                            Microsoft.Office.Interop.Excel.Range rangeMethod = wsParameters.get_Range("C7");

                            Microsoft.Office.Interop.Excel.Range rangeDataStartRow = wsParameters.get_Range("C11");
                            Microsoft.Office.Interop.Excel.Range rangeDataStartColumn = wsParameters.get_Range("C10");



                            string strStartRow = rangeDataStartRow.Value2.ToString();
                            string strStartColumn = rangeDataStartColumn.Value2.ToString();

                            int.TryParse(strStartRow, out iDataVerticalStartPosition);
                            int.TryParse(strStartColumn, out iDataHorizontalStartPosition);

                            // Microsoft.Office.Interop.Excel.Range rangeInputClassName = wsData.get_Range("D" + (iDataVerticalStartPosition - 3));

                            string strURI = rangeWebService.Value;
                            string strUsername = rangeUsername.Value;
                            string strPassword = rangePassword.Value;
                            string strMethod = rangeMethod.Value;
                            // string strInputClass = rangeInputClassName.Value;

                            List<Parameter> lsParameters = getParameterColumns(wsData, iDataHorizontalStartPosition, iDataVerticalStartPosition);

                            // generateSpreadsheet(strURI, strUsername, strPassword, strMethod, workbook);

                            compiledAssembly compiledAsm = generateCompiledAssembly(strURI, strUsername, strPassword, strMethod);

                            //if (null != (compilerResults = compiledAsm.compilerResults))

                            callWebService(generateCompiledAssembly(strURI, strUsername, strPassword, strMethod), strMethod, lsParameters, wsData, iDataVerticalStartPosition);

                        }
                        else
                        {
                            Console.WriteLine("Failed to retrieve Parameters sheet");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Failed to open " + ofdOpenFileDialog.FileName);
                    }
                    excel.Quit();
                }
            }
        }
    }


    
    public class Parameter
    {
        public string Name {get;set;}
        public string Value {get;set;}
    }

    public class compiledAssembly
    {
        public CompilerResults compilerResults { get; set; }
        public object instantiatedObject { get; set; }
    }

}
