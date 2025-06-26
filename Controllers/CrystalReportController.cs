using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Configuration; // Add this!
using System.IO;
using System.Net;
using System.Net.Http;
using System.Web.Http;
public class CrystalReportController : ApiController
{
    [HttpGet]
    [Route("api/crystalreport/exporttopdf")]
    public HttpResponseMessage ExportToPdf(string reportName, string para = null, string val = null, string printNo = null)
    {
        try
        {
            // Read from Web.config
            string reportFolder = ConfigurationManager.AppSettings["ReportFolder"];
            string server = ConfigurationManager.AppSettings["ReportServer"];
            string db = ConfigurationManager.AppSettings["ReportDatabase"];
            string user = ConfigurationManager.AppSettings["ReportUser"];
            string pwd = ConfigurationManager.AppSettings["ReportPassword"];
            string rptPath = Path.Combine(reportFolder, reportName + ".rpt");
            if (!File.Exists(rptPath))
                return Request.CreateErrorResponse(HttpStatusCode.NotFound, "Report not found: " + rptPath);

            ReportDocument reportDoc = new ReportDocument();
            reportDoc.Load(rptPath);

            // Set connection info
            var connInfo = new ConnectionInfo
            {
                ServerName = server,
                DatabaseName = db,
                UserID = user,
                Password = pwd
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
                    bool paramExists = false;
                    foreach (ParameterFieldDefinition field in crParamDefinitions)
                    {
                        if (field.Name.Equals(paramName, StringComparison.OrdinalIgnoreCase))
                        {
                            paramExists = true;
                            break;
                        }
                    }
                    if (paramExists)
                    {
                        var paraValue = new ParameterDiscreteValue { Value = paramValue };
                        var currValue = crParamDefinitions[paramName].CurrentValues;
                        currValue.Add(paraValue);
                        crParamDefinitions[paramName].ApplyCurrentValues(currValue);
                    }

                }
            }

            // Handle printNo if provided and report expects it
            if (!string.IsNullOrEmpty(printNo))
            {
                var crParamDefinitions = reportDoc.DataDefinition.ParameterFields;
                bool printNoExists = false;
                foreach (ParameterFieldDefinition field in crParamDefinitions)
                {
                    if (field.Name.Equals("printNo", StringComparison.OrdinalIgnoreCase))
                    {
                        printNoExists = true;
                        break;
                    }
                }
                if (printNoExists)
                {
                    var paraValue = new ParameterDiscreteValue { Value = printNo };
                    var currValue = crParamDefinitions["printNo"].CurrentValues;
                    currValue.Add(paraValue);
                    crParamDefinitions["printNo"].ApplyCurrentValues(currValue);
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