using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using RestSharp;
using System.Data;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System.IO;
using Form = System.Windows.Forms;

namespace APIRegressionAutomation
{
    public class AutomateAPITesting
    {
        DataTable dtAPIData = new DataTable();
        private static Excel.Workbook inputBook = null;
        private static Excel.Workbook outputBook = null;

        private static Excel.Application xlApp = null;
        private static Excel.Worksheet inputSheet = null;
        private static Excel.Worksheet outputSheet = null;

        string strAPIBaseUrl = string.Empty;
        string strAPITokenUrl = string.Empty;
        string strUserName = string.Empty;
        string strPassword = string.Empty;
        string strAccessCode = string.Empty;
        string strOutputFile = string.Empty;
        //string strLocation = string.Empty;

        const string INVALID_CREDENTIALS = "Invalid credentials";
        const string SERVER_NOT_STARTED = "Server not started";
        const string INVALID_API = "Invalid API";
        const string TOKEN_NOT_FOUND = "Token not found";
        public void ReadExcelFile(string strLocation,ref Form.TextBox txtOutPut)
        {
            var _assembly = System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase;

            //strLocation = System.IO.Path.GetDirectoryName(_assembly) + @"\HackElite.xlsx";
            DateTime date = DateTime.Now;
            string datetime = string.Format("{0:00}{1:00}{2:0000}{3:00}{4:00}{5:00}", date.Day, date.Month, date.Year, date.Hour, date.Minute, date.Second);

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            inputBook = xlApp.Workbooks.Open(strLocation);
            inputSheet = (Excel.Worksheet)inputBook.Sheets[1];
            int lastRow = inputSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            string pathString = Directory.GetCurrentDirectory() + @"\output";

            System.IO.Directory.CreateDirectory(pathString);

            strOutputFile = pathString + @"\HackElite" + '_' + datetime + ".xlsx"; 

            inputBook.SaveAs(strOutputFile);

            outputBook = xlApp.Workbooks.Open(strOutputFile);
            outputSheet = (Excel.Worksheet)outputBook.Sheets[1];

            string strAPIName = string.Empty;
            string strInputParam = string.Empty;
            string strExpectedResult = string.Empty;
            string strMessage = string.Empty;
            string result = string.Empty;
            txtOutPut.Text = "Started!!";
            txtOutPut.Text += "\r\n";

            //Console.WriteLine("Started!!");
            //Console.WriteLine();

            for (int index = 1; index <= lastRow; index++)
            {
                System.Array MyValues = (System.Array)inputSheet.get_Range("A" + index.ToString(), "D" + index.ToString()).Cells.Value;
                string strFirstCellValue = Convert.ToString(MyValues.GetValue(1, 1));
                switch (strFirstCellValue)
                {
                    case "URL":
                        {
                            strAPIBaseUrl = Convert.ToString(MyValues.GetValue(1, 2));
                            break;
                        }
                    case "TokenURL":
                        {
                            strAPITokenUrl = Convert.ToString(MyValues.GetValue(1, 2));
                            break;
                        }
                    case "Username":
                        {
                            strUserName = Convert.ToString(MyValues.GetValue(1, 2));
                            break;
                        }
                    case "Password":
                        {
                            strPassword = Convert.ToString(MyValues.GetValue(1, 2));
                            break;
                        }
                    case "code":
                        {
                            strAccessCode = Convert.ToString(MyValues.GetValue(1, 2));
                            break;
                        }
                    case "GET":
                        {
                            strAPIName = Convert.ToString(MyValues.GetValue(1, 2));
                            strInputParam = Convert.ToString(MyValues.GetValue(1, 3));
                            strExpectedResult = Convert.ToString(MyValues.GetValue(1, 4));
                            strMessage = string.Empty;
                            result = GetAPIMethod(strAPIName, strInputParam, ref strMessage, strAPITokenUrl);

                            txtOutPut.Text += "API Method Name: " + strAPIName + "  " + strInputParam + " - " + result + "  " + strMessage;
                            //Console.WriteLine(strAPIName + "  " + strInputParam + " - " + result + "  " + strMessage);
                            txtOutPut.Text += "\r\n";

                            outputSheet.Cells[index, 5] = result;
                            outputSheet.Cells[index, 6] = strMessage;
                            if (string.Equals(strExpectedResult, result, StringComparison.OrdinalIgnoreCase))
                                outputSheet.Cells[index, 7] = "passed";
                            else
                                outputSheet.Cells[index, 7] = "failed";

                            
                            break;
                        }
                    default:
                        break;
                }
                if (string.Equals(strMessage, INVALID_CREDENTIALS, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(strMessage, SERVER_NOT_STARTED, StringComparison.OrdinalIgnoreCase)  ||
                    string.Equals(strMessage, TOKEN_NOT_FOUND, StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }
            }
            outputBook.Save();
            outputBook.Close(true, strOutputFile);

            //inputBook.Close(true, strLocation);
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            System.Threading.Thread.Sleep(500);
            //Console.WriteLine();
            txtOutPut.Text += "\r\n";
            txtOutPut.Text += "Finished!!";


            //Console.WriteLine("Finished!!");
            //Console.ReadLine();
        }
       

        //public void ExecuteAPIMethod(string apiMethodUrl,string inputParam)
        //{
        //    string baseUrl = "http://localhost:3000/api/department";
        //    var client = new RestClient(baseUrl);
        //    var request = new RestRequest(Method.GET);
        //    request.AddHeader("Authorization", "Bearer bprGPVC95nsmIDxxUlPX9TUtlGXwiNvz8VFgh4niJSwhREX4VBb02I901YRspUaGnxdOHN2hS60je82hfU2fu5punBonoNYVaoCyI4TN5FYV1QKm77B3I1ba4kfXaCd6ztPWY0SQUG8PewvYCvBpjtgfiCwF8xHpb1ItSrhR6jb6lC8SNNL7hx3mhiG17kshhVNcx04IfBIpnNWyo7m14QqeuhIJs0K4al6Stl872PYYA2dVrGMpqpgQ3qAYh9uV");
        //    request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
        //    request.AddParameter("SetId", "2018/2019");
        //    IRestResponse r = client.Execute(request);
        //    string tokenResponse = r.Content;
        //    //Console.WriteLine(tokenResponse);
        //    object obj = JsonConvert.DeserializeObject(tokenResponse);
        //    Console.WriteLine(obj);
        //    Console.ReadLine();

        //}
        public string GetAPIMethod(string apiMethod, string inputParameters, ref string message,string tokenURL)
        {
            string result = "failed";

            string token = GetToken(strAPIBaseUrl, strUserName, strPassword, strAccessCode, tokenURL);
            if (string.IsNullOrEmpty(token))
            {
                message = TOKEN_NOT_FOUND;
                return result;
            }
            else if (token == INVALID_CREDENTIALS)
            {
                message = INVALID_CREDENTIALS;
                return result;
            }
            else if (token == SERVER_NOT_STARTED)
            {
                message = SERVER_NOT_STARTED;
                return result;
            }
            string authorizationCode = "Bearer " + token;
            string url = strAPIBaseUrl + apiMethod;
            if (!string.IsNullOrEmpty(inputParameters))
            {
                url += "?" + inputParameters;
            }
            var client = new RestClient(url);
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", authorizationCode);
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");

            IRestResponse r = client.Execute(request);
            if (r.ContentType == "text/html; charset=utf-8")
            {
                message = INVALID_API;
                return result;
            }

            string tokenResponse = r.Content;
            JObject obj = JObject.Parse(tokenResponse);
            if(obj.ContainsKey("status"))
                result = (string)obj["status"];

            if (obj.ContainsKey("message"))
                message = (string)obj["message"];

            return result;
        }

        public string GetToken(string strUrl, string strUserName, string strPassword, string strAccesscode,string tokenURL)
        {
            string retToken = string.Empty;
            string baseUrl = strUrl + tokenURL;
            var client = new RestClient(baseUrl);
            var request = new RestRequest(Method.POST);
            string strValue = strUserName + ":" + strPassword;
            var byteArray = Encoding.ASCII.GetBytes(strValue);
            string strtemp = Convert.ToBase64String(byteArray);
            request.AddHeader("Authorization", "Basic " + Convert.ToBase64String(byteArray));
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            request.AddParameter("code", strAccesscode);
            request.AddParameter("grant_type", "authorization_code");
            request.AddParameter("Username", strUserName);
            request.AddParameter("Password", strPassword);
            IRestResponse r = client.Execute(request);
            if (r.StatusCode == System.Net.HttpStatusCode.OK)
            {
                JObject obj = JObject.Parse(r.Content);
               // if (obj.First.Con.ContainsKey("TokenValue").c)
                    retToken = (string)obj["access_token"]["TokenValue"];
            }
            else if (r.StatusCode == System.Net.HttpStatusCode.InternalServerError)
            {
                retToken = INVALID_CREDENTIALS;
            }
            else if (r.StatusCode == 0)
            {
                retToken = SERVER_NOT_STARTED;
            }
            return retToken;
        }
    }
}
