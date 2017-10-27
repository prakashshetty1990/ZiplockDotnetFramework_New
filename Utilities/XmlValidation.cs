using Microsoft.Web.XmlTransform;
using NUnit.Framework.Interfaces;
using System;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Formatters;
using System.Xml;

namespace SWAUTCSharpFramework
{
   public class XmlValidation
    {
        //public void Execute()
        //{
       
        //    HttpWebRequest request = CreateWebRequest();
        //    XmlDocument soapEnvelopeXml = new XmlDocument();
        //    soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
        //        <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
        //        <soap:Body>
        //        < GetGeoIPResponse xmlns = ""http://www.webservicex.net/"" > 
        //        < GetGeoIPResult >
        //        < ReturnCode > 1 </ ReturnCode >
        //        < IP > 192.164.94.65 </ IP >
        //        < ReturnCodeDetails > Success </ ReturnCodeDetails >
        //        < CountryName > Austria </ CountryName >
        //        < CountryCode > AUT </ CountryCode >
        //        </ GetGeoIPResult >
        //        </ GetGeoIPResponse >
        //        </soap:Body>
        //        </soap:Envelope>");

        //    using (Stream stream = request.GetRequestStream())
        //    {
        //        soapEnvelopeXml.Save(stream);
        //    }
        //    using (WebResponse response = request.GetResponse())
        //    {
        //        using (StreamReader rd = new StreamReader(response.GetResponseStream()))
        //        {
        //            string soapResult = rd.ReadToEnd();
        //            Console.WriteLine(soapResult);
        //        }
        //    }
        //}
        ///// <summary>
        ///// Create a soap webrequest to [Url]
        ///// </summary>
        ///// <returns></returns>
        //public HttpWebRequest CreateWebRequest()
        //{
        //    HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(@"http://www.webservicex.net/geoipservice.asmx?WSDL");
        //    webRequest.Headers.Add(@"SOAP:Action");
        //    webRequest.ContentType = "text/xml;charset=\"utf-8\"";
        //    webRequest.Accept = "text/xml";
        //    webRequest.Method = "POST";
        //    return webRequest;
        //}

        //public static void validateXml()
        //{
        //    SoapCollateralPositionData();
        //    String strCountryCode, strCountryName, strReturnCodeDetails;
        //    try
        //    {

        //        strReturnCodeDetails = getdataCollateralPosition("ReturnCodeDetails");
        //        if (strReturnCodeDetails.Equals("Success"))
        //        {
        //            Common.testStepPassed("Got the Response with ReturnCodeDetails as Success");
        //        }
        //        else
        //        {
        //            Common.testStepFailed("Unable to get the Response with ReturnCodeDetails as Success");
        //        }

        //        strCountryCode = getdataCollateralPosition("CountryCode");
        //        //validation
        //        if (strCountryCode.Equals(Common.retrieve("CountryCode")))
        //        {
        //            Common.testStepPassed("Country Code is matching, Expected->" + Common.retrieve("CountryCode") + ", Actual is->" + strCountryCode);
        //        }
        //        else
        //        {
        //            Common.testStepFailed("Country Code is not matching");
        //        }

        //        strCountryName = getdataCollateralPosition("CountryName");
        //        if (strCountryName.Equals(Common.retrieve("CountryName")))
        //        {
        //            Common.testStepPassed("Country Name is matching, Expected->" + Common.retrieve("CountryName") + ", Actual is->" + strCountryName);
        //        }
        //        else
        //        {
        //            Common.testStepFailed("Country Name is not matching, Expected->" + Common.retrieve("CountryName") + ", Actual is->" + strCountryName);
        //        }

        //    }
        //    catch (Exception e)
        //    {
        //        Common.testStepFailed("Exception caught while Soap validation->" + e);
        //    }
        //}
    }
}
