using System;
using System.IO;
using System.Net;
using System.Diagnostics;
using System.Collections.Generic;
using RestSharp;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExpoIP
{
    public class Importer
    {
        public Uri RequestURI;

        private string SubDomain;
        private string ApiKey;
        private string Source;

        private bool sendRegistrationMail = true;

        private List<User> ToImport = new List<User>();
        private List<User> Imported = new List<User>();
        private List<User> FaildToImport = new List<User>();

        public Importer(string subDomain,string apiKey, string source = "dotnet_importer") {

            SubDomain = subDomain;
            ApiKey = apiKey;
            Source = source;

            RequestURI = new Uri($"https://{SubDomain}.expo-ip.com/api/user/registration/{ApiKey}?source={Source}");
            //RequestURI = new Uri($"https://{SubDomain}.expo-ip.com/registrieren?api_key={ApiKey}&source={Source}");

            //check no longer works | Nov-2020
            //needed a better implementation anyways
            //if (!checkConnection(RequestURI))
            //    throw new System.ApplicationException("Issue with the generated URI. Check for possible Typos or your internet connection");
        }

        private bool checkConnection(Uri requestURI) {

            var client = new RestClient(requestURI);
            client.Timeout = -1;
            //The idea is to prevent a download of the whole page just for a check.
            //However some servers might refuse this request...
            //var request = new RestRequest(Method.HEAD);
            //as of Nov-2020 expo-ip changed the webhook and method head no longer works...
            var request = new RestRequest(Method.POST);
            IRestResponse response = client.Execute(request);

            return response.StatusCode == HttpStatusCode.OK;
        }

        private bool registerUser(User user)
        {
            var client = new RestClient(RequestURI);
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);

            request.AlwaysMultipartFormData = true;
            
            /*set parameters requiered for registration*/
            request.AddParameter("email", user.Email);
            request.AddParameter("firstname", user.FirstName);
            request.AddParameter("lastname", user.LastName);

            /*check if optional parameters are set*/
            if (!String.IsNullOrWhiteSpace(user.Title)){
                request.AddParameter("User[title]", user.Title);
            }
            if (user.Salutation == 1 || user.Salutation == 2 || user.Salutation == 4) {
                request.AddParameter("User[salutation]", user.Salutation);
            }

            IRestResponse response = client.Execute(request);

            //TODO
            //create response class for expo-ip
            //parse json from response
            //delete this bs and write something better
            if (response.Content.Trim().Contains("E-Mail successfully sent"))
            {
                //user was successfully registerd
                Imported.Add(user);
                return true;
            }
            else
            {
                //something went wrong
                FaildToImport.Add(user);
                return false;
            }

        }

        public List<User> ReadExcel(string excelPath)
        {
            //create the Application object we can use in the member functions.
            Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
            _excelApp.Visible = false;

            string fileName = excelPath;

            Workbook workbook = _excelApp.Workbooks.Open(fileName,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
       
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            Range excelRange = worksheet.UsedRange;

            //get an object array of all of the cells in the worksheet (their values)
            object[,] valueArray = (object[,])excelRange.get_Value(
                        XlRangeValueDataType.xlRangeValueDefault);

            //access the cells
            //start at row 2 since the first row contains headlines
            for (int row = 2; row <= worksheet.UsedRange.Rows.Count; ++row)
            {

                //maybe there is a way to shorten this part...
                //read up on parse function
                //use parse instead of tryparse, parse returns and int as required
                int salutation;

                if (valueArray[row, 5] != null)
                {
                    int.TryParse(valueArray[row, 5].ToString().Trim(), out salutation);
                    //int.Parse(valueArray[row, 5].ToString().Trim(), out salutation);
                }
                else
                {
                    salutation = 0;
                }

                ToImport.Add(new User(
                        valueArray[row, 1] != null ? valueArray[row, 1].ToString().Trim() : "", //FirstName
                        valueArray[row, 2] != null ? valueArray[row, 2].ToString().Trim() : "", //SecondtName
                        valueArray[row, 3] != null ? valueArray[row, 3].ToString().Trim() : "", //Email
                        valueArray[row, 4] != null ? valueArray[row, 4].ToString().Trim() : "", //Title
                        salutation                                                              //Salutation
                        ));

            }

            //clean up stuffs
            workbook.Close(false, Type.Missing, Type.Missing);
            _excelApp.Quit();
            Marshal.ReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(_excelApp);

            return ToImport;
        }

        public string ImportUsers(List<User> users)
        {
            var importStartDate = DateTime.Now;

            foreach (User user in users)
            {
                //make a break every X users to not kill the expoip server
                if (Imported.Count + 1 % 100 == 0) {
                    System.Threading.Thread.Sleep(2000);
                }

                registerUser(user);
            }

            var importEndDate = DateTime.Now;

            string path = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + $@"\import-log-{importEndDate.ToString(@"yyyy-MM-dd-HH-mm-ss")}.txt";

            if (!File.Exists(path))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("Import Log");
                    sw.WriteLine($"------------------------");
                    sw.WriteLine($"Request: {RequestURI}");
                    sw.WriteLine($"Started: {importStartDate}");
                    sw.WriteLine($"Finished: {importEndDate}");
                    sw.WriteLine($"------------------------");
                    sw.WriteLine($"Users to import: {ToImport.Count}");
                    sw.WriteLine($"Users successfully imported: {Imported.Count}");
                    sw.WriteLine($"Users failed to import: {FaildToImport.Count}");
                    sw.WriteLine($"------------------------");
                    sw.WriteLine("List of Users failed to import:");

                    foreach (User user in FaildToImport)
                    {
                        sw.WriteLine($"{user.FirstName} {user.LastName} - {user.Email}");
                    }


                }
            }

            return $"Import Done\n" +
                    $"Users to import: {ToImport.Count}\n" +
                    $"Users successfully imported: {Imported.Count}\n" +
                    $"Users failed to import: {FaildToImport.Count}\n";
        }

    } /* End Class */

} /* End Namespace */
