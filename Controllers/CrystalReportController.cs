using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Http;
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

            // --- Export to PDF or Excel ---
            ExportFormatType exportFormat = ExportFormatType.PortableDocFormat;
            string contentType = "application/pdf";
            string fileExt = "pdf";

            // Detect if Excel is requested
            if (!string.IsNullOrEmpty(docType) && docType.ToLower().Contains("excel"))
            {
                exportFormat = ExportFormatType.ExcelWorkbook; // Or ExcelRecord for raw .xls
                contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                fileExt = "xlsx";
            }

            Stream exportStream = reportDoc.ExportToStream(exportFormat);
            exportStream.Seek(0, SeekOrigin.Begin); // Ensure stream is at beginning

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
        catch (Exception ex)
        {
            return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.ToString());
        }
    }

}
