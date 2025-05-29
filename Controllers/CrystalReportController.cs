
using System.Net;
using System.Net.Http;
using System.Web.Http;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.IO;
using System;

public class CrystalReportController : ApiController
{
    [HttpGet]
    [Route("api/crystalreport/exporttopdf")]
    public HttpResponseMessage ExportToPdf(string reportName, string para = null, string val = null)
    {
        try
        {
            string reportFolder = @"D:\Project\Reports";
            string rptPath = Path.Combine(reportFolder, reportName + ".rpt");
            if (!File.Exists(rptPath))
                return Request.CreateErrorResponse(HttpStatusCode.NotFound, "Report not found: " + rptPath);

            ReportDocument reportDoc = new ReportDocument();
            reportDoc.Load(rptPath);

            // Set connection info as needed
            var connInfo = new ConnectionInfo
            {
                ServerName = "10.2.80.239",
                DatabaseName = "BCS",
                UserID = "sa",
                Password = "2770"
            };
            foreach (Table table in reportDoc.Database.Tables)
            {
                var logonInfo = table.LogOnInfo;
                logonInfo.ConnectionInfo = connInfo;
                table.ApplyLogOnInfo(logonInfo);
            }

            // Set parameters if any
            if (!string.IsNullOrEmpty(para) && !string.IsNullOrEmpty(val))
            {
                // Replace underscores with ampersands in VALUES
                string safeVal = val.Replace("_", "&");
                string[] paraArr = para.Split('|');
                string[] valArr = safeVal.Split('|');
                if (paraArr.Length != valArr.Length)
                    throw new Exception("Number of parameters does not match values. para: " + para + " val: " + val);

                var crParamDefinitions = reportDoc.DataDefinition.ParameterFields;
                for (int i = 0; i < paraArr.Length; i++)
                {
                    var paramName = paraArr[i];
                    var paramValue = valArr[i];
                    var paraValue = new ParameterDiscreteValue { Value = paramValue };
                    var currValue = crParamDefinitions[paramName].CurrentValues;
                    currValue.Add(paraValue);
                    crParamDefinitions[paramName].ApplyCurrentValues(currValue);
                }
            }

            // Export to PDF and return as stream
            var stream = reportDoc.ExportToStream(ExportFormatType.PortableDocFormat);
            var result = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(stream)
            };
            result.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/pdf");
            result.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
            {
                FileName = $"{reportName}.pdf"
            };
            return result;
        }
        catch (Exception ex)
        {
            return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.ToString());
        }
    }

}
