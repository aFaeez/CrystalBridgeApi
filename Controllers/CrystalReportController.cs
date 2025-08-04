using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;

public class CrystalReportController : ApiController
{
    [HttpGet]
    [Route("api/crystalreport/exporttopdf")]
    public HttpResponseMessage ExportToPdf(
        string reportName,
        string para = null,
        string val = null,
        string printNo = null,
        string sDate = null,
        string eDate = null,
        string sFormula = null,
        string pFormula = null,
        string isPW = null,
        string isQSME = null,
        string file = null,
        string docType = null
        )
    {
        try
        {
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

            // Set DB login info
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

            var crParamDefinitions = reportDoc.DataDefinition.ParameterFields;

            // --- 1. Set generic parameters via para/val if provided ---
            if (!string.IsNullOrEmpty(para) && !string.IsNullOrEmpty(val))
            {
                string safeVal = val.Replace("_", "&");
                string[] paraArr = para.Split('|');
                string[] valArr = safeVal.Split('|');
                if (paraArr.Length != valArr.Length)
                    throw new Exception("Number of parameters does not match values. para: " + para + " val: " + val);

                for (int i = 0; i < paraArr.Length; i++)
                {
                    var paramName = paraArr[i];
                    var paramValue = valArr[i];
                    foreach (ParameterFieldDefinition field in crParamDefinitions)
                    {
                        if (field.Name.Equals(paramName, StringComparison.OrdinalIgnoreCase))
                        {
                            var paraValue = new ParameterDiscreteValue { Value = paramValue };
                            var currValue = crParamDefinitions[paramName].CurrentValues;
                            currValue.Add(paraValue);
                            crParamDefinitions[paramName].ApplyCurrentValues(currValue);
                            break;
                        }
                    }
                }
            }

            // ---  Set printNo if provided ---
            if (!string.IsNullOrEmpty(printNo))
            {
                foreach (ParameterFieldDefinition field in crParamDefinitions)
                {
                    if (field.Name.Equals("printNo", StringComparison.OrdinalIgnoreCase))
                    {
                        var paraValue = new ParameterDiscreteValue { Value = printNo };
                        var currValue = crParamDefinitions["printNo"].CurrentValues;
                        currValue.Add(paraValue);
                        crParamDefinitions["printNo"].ApplyCurrentValues(currValue);
                        break;
                    }
                }
            }

            // --- Set sDate/eDate if provided, and convert to DateTime (for reports expecting Date/DateTime) ---
            var culture = new System.Globalization.CultureInfo("en-GB");

            if (!string.IsNullOrEmpty(sDate))
            {
                foreach (ParameterFieldDefinition field in crParamDefinitions)
                {
                    if (field.Name.Equals("sDate", StringComparison.OrdinalIgnoreCase) || field.Name.Equals("@sDate", StringComparison.OrdinalIgnoreCase))
                    {
                        DateTime dateValue;
                        if (DateTime.TryParse(sDate, culture, System.Globalization.DateTimeStyles.None, out dateValue))
                        {
                            var paraValue = new ParameterDiscreteValue { Value = dateValue };
                            var currValue = crParamDefinitions[field.Name].CurrentValues;
                            currValue.Add(paraValue);
                            crParamDefinitions[field.Name].ApplyCurrentValues(currValue);
                        }
                        else
                        {
                            // fallback: set as string
                            var paraValue = new ParameterDiscreteValue { Value = sDate };
                            var currValue = crParamDefinitions[field.Name].CurrentValues;
                            currValue.Add(paraValue);
                            crParamDefinitions[field.Name].ApplyCurrentValues(currValue);
                        }
                        break;
                    }
                }
            }
            if (!string.IsNullOrEmpty(eDate))
            {
                foreach (ParameterFieldDefinition field in crParamDefinitions)
                {
                    if (field.Name.Equals("eDate", StringComparison.OrdinalIgnoreCase) || field.Name.Equals("@eDate", StringComparison.OrdinalIgnoreCase))
                    {
                        DateTime dateValue;
                        if (DateTime.TryParse(eDate, culture, System.Globalization.DateTimeStyles.None, out dateValue))
                        {
                            var paraValue = new ParameterDiscreteValue { Value = dateValue };
                            var currValue = crParamDefinitions[field.Name].CurrentValues;
                            currValue.Add(paraValue);
                            crParamDefinitions[field.Name].ApplyCurrentValues(currValue);
                        }
                        else
                        {
                            // fallback: set as string
                            var paraValue = new ParameterDiscreteValue { Value = eDate };
                            var currValue = crParamDefinitions[field.Name].CurrentValues;
                            currValue.Add(paraValue);
                            crParamDefinitions[field.Name].ApplyCurrentValues(currValue);
                        }
                        break;
                    }
                }
            }
            // --- Apply custom Record Selection Formula based on sFormula and pFormula ---
            if (!string.IsNullOrEmpty(sFormula) && !string.IsNullOrEmpty(pFormula))
            {
                var fieldNames = pFormula.Contains("|") ? pFormula.Split('|') : new string[] { pFormula };
                var fieldValues = sFormula.Contains("|") ? sFormula.Split('|') : new string[] { sFormula };

                if (fieldNames.Length == fieldValues.Length)
                {
                    var formulaBuilder = new List<string>();

                    for (int i = 0; i < fieldNames.Length; i++)
                    {
                        var rawValue = fieldValues[i]
                            .Replace("**", "','") // handles multiple values
                            .Replace("*", "");    // remove single *

                        var condition = $"{fieldNames[i].Trim()} in ['{rawValue}']";
                        formulaBuilder.Add(condition);
                    }

                    string finalFormula = string.Join(" and ", formulaBuilder);

                    // Apply to the report
                    reportDoc.RecordSelectionFormula = finalFormula;
                }
            }


            // --- isPW ---
            if (!string.IsNullOrEmpty(isPW))
            {
                foreach (ParameterFieldDefinition field in crParamDefinitions)
                {
                    if (field.Name.Equals("isPW", StringComparison.OrdinalIgnoreCase))
                    {
                        var paraValue = new ParameterDiscreteValue { Value = isPW };
                        var currValue = crParamDefinitions[field.Name].CurrentValues;
                        currValue.Add(paraValue);
                        crParamDefinitions[field.Name].ApplyCurrentValues(currValue);
                        break;
                    }
                }
            }

            // --- isQSME ---
            if (!string.IsNullOrEmpty(isQSME))
            {
                foreach (ParameterFieldDefinition field in crParamDefinitions)
                {
                    if (field.Name.Equals("isQSME", StringComparison.OrdinalIgnoreCase))
                    {
                        var paraValue = new ParameterDiscreteValue { Value = isQSME };
                        var currValue = crParamDefinitions[field.Name].CurrentValues;
                        currValue.Add(paraValue);
                        crParamDefinitions[field.Name].ApplyCurrentValues(currValue);
                        break;
                    }
                }
            }

            // --- file ---
            if (!string.IsNullOrEmpty(file))
            {
                foreach (ParameterFieldDefinition field in crParamDefinitions)
                {
                    if (field.Name.Equals("file", StringComparison.OrdinalIgnoreCase))
                    {
                        var paraValue = new ParameterDiscreteValue { Value = file };
                        var currValue = crParamDefinitions[field.Name].CurrentValues;
                        currValue.Add(paraValue);
                        crParamDefinitions[field.Name].ApplyCurrentValues(currValue);
                        break;
                    }
                }
            }

            // --- docType ---
            if (!string.IsNullOrEmpty(docType))
            {
                foreach (ParameterFieldDefinition field in crParamDefinitions)
                {
                    if (field.Name.Equals("docType", StringComparison.OrdinalIgnoreCase))
                    {
                        var paraValue = new ParameterDiscreteValue { Value = docType };
                        var currValue = crParamDefinitions[field.Name].CurrentValues;
                        currValue.Add(paraValue);
                        crParamDefinitions[field.Name].ApplyCurrentValues(currValue);
                        break;
                    }
                }
            }

            // --- Determine export format ---
            ExportFormatType exportFormat = ExportFormatType.PortableDocFormat;
            string contentType = "application/pdf";
            string fileExt = "pdf";

            if (!string.IsNullOrEmpty(docType) && docType.ToLower().Contains("excel"))
            {
                exportFormat = ExportFormatType.ExcelWorkbook;
                contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                fileExt = "xlsx";
            }

            // --- Special logic only for 3 specific report names ---
            bool isCertReport =
                reportName.Equals("CertificateWithAppendixWithFS_IN", StringComparison.OrdinalIgnoreCase) ||
                reportName.Equals("CertificateWithAppendixWithFS_EXT", StringComparison.OrdinalIgnoreCase) ||
                reportName.Equals("CertificateWithAppendix", StringComparison.OrdinalIgnoreCase);

            if (isCertReport)
            {
                // 1. Export to temp path for download
                string tempPath = Path.Combine(
                    ConfigurationManager.AppSettings["ReportFolder"],
                    $"{Guid.NewGuid()}.{fileExt}"
                );
                reportDoc.ExportToDisk(exportFormat, tempPath);

                // 2. Export copy to certificate folder
                string certPath = ConfigurationManager.AppSettings["SPSLoadCertificatePath"];
                string safeVal = val?.Replace("/", "").Replace("|", "_").Replace("&", "_") ?? "Default";
                string printSuffix = string.IsNullOrEmpty(printNo) ? "" : $"_{printNo}";
                string certFileName = Path.Combine(certPath, $"{safeVal}{printSuffix}.pdf");
                reportDoc.ExportToDisk(ExportFormatType.PortableDocFormat, certFileName);

                // 3. Return stream of the temp file
                var result = new HttpResponseMessage(HttpStatusCode.OK);
                result.Content = new StreamContent(new FileStream(tempPath, FileMode.Open, FileAccess.Read));
                result.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(contentType);
                result.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
                {
                    FileName = $"{reportName}.{fileExt}"
                };

                // 4. Delete temp file after sending
                _ = Task.Run(() =>
                {
                    try
                    {
                        System.Threading.Thread.Sleep(5000);
                        if (File.Exists(tempPath))
                            File.Delete(tempPath);
                    }
                    catch { /* ignore */ }
                });

                return result;
            }
            else
            {
                // --- Normal path: export to memory stream and return ---
                Stream exportStream = reportDoc.ExportToStream(exportFormat);
                exportStream.Seek(0, SeekOrigin.Begin);

                var result = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StreamContent(exportStream)
                };
                result.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(contentType);
                result.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
                {
                    FileName = $"{reportName}.{fileExt}"
                };
                return result;
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine("ExportToPdf Error: " + ex.ToString());
            return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.ToString());
        }
    }

}
