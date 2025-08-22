using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Hosting;

public class CrystalReportController : ApiController
{
    private static void ApplyDbLogon(
        ReportDocument reportDoc,
        string provider,     // "oledb" or "odbc" (RDO)
        string serverOrDsn,  // OLE DB server OR ODBC DSN name
        string database,
        string user,
        string pwd,
        bool isOdbcRdo = false)
    {
        var conn = new ConnectionInfo
        {
            ServerName = serverOrDsn,
            UserID = user,
            Password = pwd,
            IntegratedSecurity = false
        };

        // RDO quirk: do NOT push DatabaseName, and do NOT force <db>.dbo.<table>
        if (!isOdbcRdo)
            conn.DatabaseName = database ?? "";

        void ApplyToTable(Table table)
        {
            var logon = table.LogOnInfo;
            logon.ConnectionInfo = conn;
            table.ApplyLogOnInfo(logon);

            // Location normalization:
            // - OLE DB: qualify if bare
            // - ODBC RDO: leave as-is; if completely bare, prefer "dbo.Table"
            var loc = table.Location ?? "";
            int semi = loc.IndexOf(';'); // C# 7.3-compatible (no range operator)
            var locNoSuffix = semi >= 0 ? loc.Substring(0, semi) : loc;

            if (!isOdbcRdo)
            {
                // OLE DB path: qualify to <db>.dbo.<table> if needed
                if (!string.IsNullOrEmpty(database) &&
                    (!locNoSuffix.Contains(".") ||
                     string.Equals(locNoSuffix, table.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    try { table.Location = string.Format("{0}.dbo.{1}", database, table.Name); } catch { }
                }
            }
            else
            {
                // ODBC (RDO) path: DO NOT force database; keep schema-only
                if (!locNoSuffix.Contains(".") ||
                    string.Equals(locNoSuffix, table.Name, StringComparison.OrdinalIgnoreCase))
                {
                    try { table.Location = string.Format("dbo.{0}", table.Name); } catch { }
                }
            }
        }

        foreach (Table t in reportDoc.Database.Tables) ApplyToTable(t);
        foreach (ReportDocument sub in reportDoc.Subreports)
            foreach (Table t in sub.Database.Tables) ApplyToTable(t);

        // Inform the report as well
        if (!isOdbcRdo)
            reportDoc.SetDatabaseLogon(user, pwd, serverOrDsn, database);
        else
        {
            // RDO: avoid passing database at report level
            try { reportDoc.SetDatabaseLogon(user, pwd); } catch { }

            // Some RDO reports must have DataSourceConnections SetConnection called.
            // We do a 2-pass approach: (1) DSN only, (2) DSN + Database if needed.

            // Pass 1: DSN only
            bool needPass2 = false;
            try
            {
                for (int i = 0; i < reportDoc.DataSourceConnections.Count; i++)
                {
                    var c = reportDoc.DataSourceConnections[i];
                    try { c.SetConnection(serverOrDsn, "", false); } catch { needPass2 = true; }
                    try { c.SetLogon(user, pwd); } catch { needPass2 = true; }
                }
            }
            catch { needPass2 = true; }

            // Quick probe: if connectivity still fails for any table, try pass 2
            try
            {
                foreach (Table t in reportDoc.Database.Tables)
                {
                    try { if (!t.TestConnectivity()) { needPass2 = true; break; } }
                    catch { needPass2 = true; break; }
                }
                if (!needPass2)
                {
                    foreach (ReportDocument sub in reportDoc.Subreports)
                    {
                        foreach (Table t in sub.Database.Tables)
                        {
                            try { if (!t.TestConnectivity()) { needPass2 = true; break; } }
                            catch { needPass2 = true; break; }
                        }
                        if (needPass2) break;
                    }
                }
            }
            catch { needPass2 = true; }

            // Pass 2: DSN + Database (some RDO reports need the catalog here)
            if (needPass2)
            {
                try
                {
                    for (int i = 0; i < reportDoc.DataSourceConnections.Count; i++)
                    {
                        var c = reportDoc.DataSourceConnections[i];
                        try { c.SetConnection(serverOrDsn, database ?? "", false); } catch { }
                        try { c.SetLogon(user, pwd); } catch { }
                    }
                }
                catch { }
            }
        }
        try { reportDoc.VerifyDatabase(); } catch { /* non-fatal */ }
    }


    [HttpGet]
    [Route("api/test/hello")]
    public IHttpActionResult Hello()
    {
        return Ok("Hello World! API is working.");
    }

    private static string TestConnections(ReportDocument doc)
    {
        try
        {
            // Main report
            foreach (Table t in doc.Database.Tables)
            {
                try
                {
                    if (!t.TestConnectivity())
                        return string.Format("Connectivity failed: main table '{0}' (Location='{1}')", t.Name, t.Location);
                }
                catch (Exception ex)
                {
                    return string.Format("Connectivity EXCEPTION: main table '{0}' (Location='{1}') -> {2}: {3}",
                        t.Name, t.Location, ex.GetType().Name, ex.Message);
                }
            }

            // Subreports
            foreach (ReportDocument sub in doc.Subreports)
            {
                foreach (Table t in sub.Database.Tables)
                {
                    try
                    {
                        if (!t.TestConnectivity())
                            return string.Format("Connectivity failed: subreport '{0}' table '{1}' (Location='{2}')",
                                sub.Name, t.Name, t.Location);
                    }
                    catch (Exception ex)
                    {
                        return string.Format("Connectivity EXCEPTION: subreport '{0}' table '{1}' (Location='{2}') -> {3}: {4}",
                            sub.Name, t.Name, t.Location, ex.GetType().Name, ex.Message);
                    }
                }
            }

            // DataSourceConnections details (sometimes 0 or more)
            try
            {
                for (int i = 0; i < doc.DataSourceConnections.Count; i++)
                {
                    var c = doc.DataSourceConnections[i];
                    // Just touching it can throw if bad
                    var serverName = c.ServerName ?? "";
                    var dbName = c.DatabaseName ?? "";
                }
            }
            catch (Exception ex)
            {
                return string.Format("DataSourceConnections error -> {0}: {1}", ex.GetType().Name, ex.Message);
            }

            return null; // all good
        }
        catch (Exception ex)
        {
            return string.Format("TestConnections wrapper error -> {0}: {1}", ex.GetType().Name, ex.Message);
        }
    }

    private static void SetDiscreteParam(ParameterFieldDefinition field, object value)
    {
        if (!string.IsNullOrEmpty(field.ReportName)) return; // skip linked parameter
        var vals = new ParameterValues();
        vals.Clear(); // ensure no leftover defaults
        vals.Add(new ParameterDiscreteValue { Value = value });
        try { field.ApplyCurrentValues(vals); } catch { /* ignore */ }
    }

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
        string docType = null,
        string userName = null,
        string connType = null)
    {
        ReportDocument reportDoc = null;

        try
        {
            // --- Config ---
            string reportFolder = ConfigurationManager.AppSettings["ReportFolder"];
            if (reportFolder.StartsWith("~"))
            {
                reportFolder = HostingEnvironment.MapPath(reportFolder);
            }

            string server = ConfigurationManager.AppSettings["ReportServer"];        // OLE DB host
            string db = ConfigurationManager.AppSettings["ReportDatabase"];
            string user = ConfigurationManager.AppSettings["ReportUser"];
            string pwd = ConfigurationManager.AppSettings["ReportPassword"];
            string odbcDsn = ConfigurationManager.AppSettings["ReportOdbcDsn"];       // ODBC DSN
            string odbcUser = ConfigurationManager.AppSettings["ReportOdbcUser"];
            string odbcPwd = ConfigurationManager.AppSettings["ReportOdbcPassword"];
            string defaultProvider = (ConfigurationManager.AppSettings["ReportDbProvider"] ?? "oledb").ToLowerInvariant();
            string provider = string.IsNullOrWhiteSpace(connType) ? defaultProvider : connType.ToLowerInvariant();

            // --- Locate report ---
            string rptPath = Path.Combine(reportFolder, reportName + ".rpt");
            if (!File.Exists(rptPath))
                return Request.CreateErrorResponse(HttpStatusCode.NotFound, "Report not found: " + rptPath);

            // --- Load report ---
            reportDoc = new ReportDocument();
            reportDoc.Load(rptPath);

            // --- Apply DB logon (main + subreports) ---
            if (provider == "odbc")
            {
                var dsn = string.IsNullOrWhiteSpace(odbcDsn) ? server : odbcDsn;
                ApplyDbLogon(reportDoc, "odbc", dsn, db, odbcUser, odbcPwd, isOdbcRdo: true);
            }
            else
            {
                ApplyDbLogon(reportDoc, "oledb", server, db, user, pwd, isOdbcRdo: false);
            }

            var connError = TestConnections(reportDoc);
            if (!string.IsNullOrEmpty(connError))
            {
                // Return detailed reason instead of generic COMException
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, connError);
            }

            var crParams = reportDoc.DataDefinition.ParameterFields;

            // --- 1) Generic parameters via para/val ---
            if (!string.IsNullOrEmpty(para) && !string.IsNullOrEmpty(val))
            {
                string safeVal = val.Replace("_", "&");
                string[] paraArr = para.Split('|');
                string[] valArr = safeVal.Split('|');
                if (paraArr.Length != valArr.Length)
                    throw new Exception("Number of parameters does not match values. para: " + para + " val: " + val);

                for (int i = 0; i < paraArr.Length; i++)
                {
                    string paramName = paraArr[i];
                    string paramValue = valArr[i];

                    foreach (ParameterFieldDefinition f in crParams)
                    {
                        if (f.Name.Equals(paramName, StringComparison.OrdinalIgnoreCase))
                        {
                            SetDiscreteParam(f, paramValue);
                            break;
                        }
                    }
                }
            }

            // --- 2) printNo ---
            if (!string.IsNullOrEmpty(printNo))
            {
                foreach (ParameterFieldDefinition f in crParams)
                    if (f.Name.Equals("printNo", StringComparison.OrdinalIgnoreCase))
                    { SetDiscreteParam(f, printNo); break; }
            }

            // --- 3) sDate / eDate ---
            var culture = new CultureInfo("en-GB");

            if (!string.IsNullOrEmpty(sDate))
            {
                foreach (ParameterFieldDefinition f in crParams)
                {
                    if (f.Name.Equals("sDate", StringComparison.OrdinalIgnoreCase) ||
                        f.Name.Equals("@sDate", StringComparison.OrdinalIgnoreCase))
                    {
                        DateTime dt;
                        if (DateTime.TryParse(sDate, culture, DateTimeStyles.None, out dt))
                            SetDiscreteParam(f, dt);
                        else
                            SetDiscreteParam(f, sDate);
                        break;
                    }
                }
            }

            if (!string.IsNullOrEmpty(eDate))
            {
                foreach (ParameterFieldDefinition f in crParams)
                {
                    if (f.Name.Equals("eDate", StringComparison.OrdinalIgnoreCase) ||
                        f.Name.Equals("@eDate", StringComparison.OrdinalIgnoreCase))
                    {
                        DateTime dt;
                        if (DateTime.TryParse(eDate, culture, DateTimeStyles.None, out dt))
                            SetDiscreteParam(f, dt);
                        else
                            SetDiscreteParam(f, eDate);
                        break;
                    }
                }
            }

            // --- 4) Record Selection Formula via sFormula/pFormula ---
            if (!string.IsNullOrEmpty(sFormula) && !string.IsNullOrEmpty(pFormula))
            {
                var fieldNames = pFormula.Contains("|") ? pFormula.Split('|') : new[] { pFormula };
                var fieldVals = sFormula.Contains("|") ? sFormula.Split('|') : new[] { sFormula };

                if (fieldNames.Length == fieldVals.Length)
                {
                    var clauses = new List<string>();
                    for (int i = 0; i < fieldNames.Length; i++)
                    {
                        var raw = fieldVals[i]
                            .Replace("**", "','") // multi-values separator
                            .Replace("*", "");    // strip single *
                        clauses.Add(string.Format("{0} in ['{1}']", fieldNames[i].Trim(), raw));
                    }
                    string finalFormula = string.Join(" and ", clauses);
                    reportDoc.RecordSelectionFormula = finalFormula;
                }
            }

            // --- 5) Flags (isPW, isQSME, file, docType) ---
            if (!string.IsNullOrEmpty(isPW))
                foreach (ParameterFieldDefinition f in crParams)
                    if (f.Name.Equals("isPW", StringComparison.OrdinalIgnoreCase)) { SetDiscreteParam(f, isPW); break; }

            if (!string.IsNullOrEmpty(isQSME))
                foreach (ParameterFieldDefinition f in crParams)
                    if (f.Name.Equals("isQSME", StringComparison.OrdinalIgnoreCase)) { SetDiscreteParam(f, isQSME); break; }

            if (!string.IsNullOrEmpty(file))
                foreach (ParameterFieldDefinition f in crParams)
                    if (f.Name.Equals("file", StringComparison.OrdinalIgnoreCase)) { SetDiscreteParam(f, file); break; }

            if (!string.IsNullOrEmpty(docType))
                foreach (ParameterFieldDefinition f in crParams)
                    if (f.Name.Equals("docType", StringComparison.OrdinalIgnoreCase)) { SetDiscreteParam(f, docType); break; }

            // --- 6) Export format ---
            ExportFormatType exportFormat = ExportFormatType.PortableDocFormat;
            string contentType = "application/pdf";
            string fileExt = "pdf";

            if (!string.IsNullOrEmpty(docType) && docType.ToLower().Contains("excel"))
            {
                exportFormat = ExportFormatType.ExcelWorkbook;
                contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                fileExt = "xlsx";
            }

            // --- 7) Special handling for 3 certificate reports ---
            bool isCertReport =
                reportName.Equals("CertificateWithAppendixWithFS_IN", StringComparison.OrdinalIgnoreCase) ||
                reportName.Equals("CertificateWithAppendixWithFS_EXT", StringComparison.OrdinalIgnoreCase) ||
                reportName.Equals("CertificateWithAppendix", StringComparison.OrdinalIgnoreCase);

            if (isCertReport)
            {
                if (!string.IsNullOrWhiteSpace(userName))
                {
                    foreach (ParameterFieldDefinition f in crParams)
                        if (f.Name.Equals("userName", StringComparison.OrdinalIgnoreCase))
                        { SetDiscreteParam(f, userName); break; }
                }

                string tempFile = Path.Combine(reportFolder, string.Format("{0}.{1}", Guid.NewGuid(), fileExt));
                reportDoc.ExportToDisk(exportFormat, tempFile);

                string certPath = ConfigurationManager.AppSettings["SPSLoadCertificatePath"];
                string safeFile = val != null ? val.Replace("/", "").Replace("|", "_").Replace("&", "_") : "Cert";
                string suffix = string.IsNullOrEmpty(printNo) ? "" : "_" + printNo;
                string certFile = Path.Combine(certPath, string.Format("{0}{1}.pdf", safeFile, suffix));

                // Always export a PDF copy to the cert path
                reportDoc.ExportToDisk(ExportFormatType.PortableDocFormat, certFile);

                var result = new HttpResponseMessage(HttpStatusCode.OK);
                result.Content = new StreamContent(new FileStream(tempFile, FileMode.Open, FileAccess.Read));
                result.Content.Headers.ContentType =
                    new System.Net.Http.Headers.MediaTypeHeaderValue(contentType);
                result.Content.Headers.ContentDisposition =
                    new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
                    { FileName = string.Format("{0}.{1}", reportName, fileExt) };

                _ = System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        System.Threading.Thread.Sleep(5000);
                        if (File.Exists(tempFile)) File.Delete(tempFile);
                    }
                    catch { }
                });

                return result;
            }

            // --- 8) Normal path: export to stream ---
            Stream exportStream = reportDoc.ExportToStream(exportFormat);
            exportStream.Seek(0, SeekOrigin.Begin);

            var ok = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(exportStream)
            };
            ok.Content.Headers.ContentType =
                new System.Net.Http.Headers.MediaTypeHeaderValue(contentType);
            ok.Content.Headers.ContentDisposition =
                new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
                { FileName = string.Format("{0}.{1}", reportName, fileExt) };

            return ok;
        }
        catch (Exception ex)
        {
            Console.WriteLine("ExportToPdf Error: " + ex);
            return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.ToString());
        }
        finally
        {
            // Important: prevent file/handle leaks
            try
            {
                if (reportDoc != null)
                {
                    reportDoc.Close();
                    reportDoc.Dispose();
                }
            }
            catch { }
        }
    }
}
