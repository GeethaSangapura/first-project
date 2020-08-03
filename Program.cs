using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net;
using System.IO;
using Microsoft.SharePoint.Client;
using System.Xml;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using Wictor.Office365;
using System.Web.Services;
using System.Web.Services.Protocols;
using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client.Utilities;
using System.Globalization;
using System.Configuration;
using PSLibrary = Microsoft.Office.Project.Server.Library;
using System.Security;
using System.Net.Mail;
using System.Threading;
using Newtonsoft.Json.Linq;

namespace e2eSendRMOMail
{
    class Program
    {
        public static string URL = System.Configuration.ConfigurationSettings.AppSettings["URL"];
        
        public static string UserName = System.Configuration.ConfigurationSettings.AppSettings["UserName"];
        public static string Password = System.Configuration.ConfigurationSettings.AppSettings["Password"];
        private static string ServerName = System.Configuration.ConfigurationSettings.AppSettings["ServerName"];
        private static string DatabaseName = System.Configuration.ConfigurationSettings.AppSettings["DatabaseName"];
        private static string DbUserName = System.Configuration.ConfigurationSettings.AppSettings["DbUserName"];
        private static string DbPassword = System.Configuration.ConfigurationSettings.AppSettings["DbPassword"];
        public static string RedirectURL = System.Configuration.ConfigurationSettings.AppSettings["RedirectURL"];
        public static string MyURL = System.Configuration.ConfigurationSettings.AppSettings["MyURL"];

        public static string ReleaseRedirectURL = System.Configuration.ConfigurationSettings.AppSettings["ReleaseRedirectURL"];
        public static string serviceURL = System.Configuration.ConfigurationSettings.AppSettings["ServiceURL"];
        public static string ErrorFlag = "";
        public static string ResourceEmailValue = "";

        private static SqlConnection con = new SqlConnection("Data Source=" + ServerName + ";Database=" + DatabaseName + "; User Id=" + DbUserName + "; password= " + DbPassword + "");

        static void Main(string[] args)
        {
            try
            {
                getItemDetails();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadKey();
            }
        }


        public static void getItemDetails()
        {
            try
            {
                string TrackingId = "", EventId = "", To = "", Cc = ""; string ProjectManagerMail = "";
                MsOnlineClaimsHelper claimsHelper = new MsOnlineClaimsHelper(URL, UserName, Password);
                // using (ClientContext ctx = new ClientContext(URL))
                 
                using (ClientContext ctx = GetContext())
                {
                    
                    //  ctx.ExecutingWebRequest += claimsHelper.clientContext_ExecutingWebRequest;
                    Web _oweb = ctx.Web;

                    List _olist = _oweb.Lists.GetByTitle("BconeEmailData");


                    //working
                   // var CamlQuery = new CamlQuery() { ViewXml = "<View><Query><Where><Or><Eq><FieldRef Name='Flag' /><Value Type='Text'>5</Value></Eq><Or><Eq><FieldRef Name='Flag' /><Value Type='Text'>6</Value></Eq><Or><Eq><FieldRef Name='Flag' /><Value Type='Text'>7</Value></Eq><Or><Eq><FieldRef Name='Flag' /><Value Type='Text'>9</Value></Eq><Or><Eq><FieldRef Name='Flag' /><Value Type='Text'>10</Value></Eq><Or><Eq><FieldRef Name='Flag' /><Value Type='Text'>40</Value></Eq><Eq><FieldRef Name='Flag' /><Value Type='Text'>41</Value></Eq></Or></Or></Or></Or></Or></Or></Where></Query></View>" };
                    var CamlQuery = new CamlQuery() { ViewXml = "<View><Query><Where><Eq><FieldRef Name='EventId'/><Value Type='Text'>25</Value></Eq></Where></Query></View>" };
                    //var CamlQuery = new CamlQuery() { ViewXml = "<View><Query><Where><Eq><FieldRef Name='Flag'/><Value Type='Text'>10</Value></Eq></Where></Query></View>" };
                    ListItemCollection _olistItemsCollection = _olist.GetItems(CamlQuery);
                    
                    ctx.Load(_olistItemsCollection);
                    ctx.ExecuteQuery();
                    Console.WriteLine("Process Started");
                    foreach (ListItem items in _olistItemsCollection)
                    {

                        TrackingId = Convert.ToString(items["TrackingId"]);
                        EventId = Convert.ToString(items["EventId"]);
                         To = Convert.ToString(items["To"]);
                        Cc = Convert.ToString(items["Cc"]);
                         
                        if (To != "" & To != "Na")
                        {

                            if (items["Flag"].ToString() == "7")
                            {

                                SendDynamicTableEmailRelease(ctx, _oweb, EventId, To, Cc, TrackingId);

                            }
                            else if (items["Flag"].ToString() == "9")
                            {
                                List _olistRRFMaster = _oweb.Lists.GetByTitle("ResourceAllocationDetails");
                                var CamlQueryRRF = new CamlQuery() { ViewXml = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>" + TrackingId + "</Value></Eq></Where></Query></View>" };
                                ListItemCollection _olistItemsRRFCollection = _olistRRFMaster.GetItems(CamlQueryRRF);
                                ctx.Load(_olistItemsRRFCollection);
                                ctx.ExecuteQuery();
                                foreach (ListItem itemsRRF in _olistItemsRRFCollection)
                                {
                                    sendEmailRelease(ctx, _oweb, itemsRRF, EventId, To, Cc);
                                }

                            }
                            else if (items["Flag"].ToString() == "40" || items["Flag"].ToString() == "41") //41 extenion reject anf 40 early release rject
                            {
                                List _olistRRFMaster = _oweb.Lists.GetByTitle("ResourceAllocationDetails");
                                var CamlQueryRRF = new CamlQuery() { ViewXml = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>" + TrackingId + "</Value></Eq></Where></Query></View>" };
                                ListItemCollection _olistItemsRRFCollection = _olistRRFMaster.GetItems(CamlQueryRRF);
                                ctx.Load(_olistItemsRRFCollection);
                                ctx.ExecuteQuery();
                                foreach (ListItem itemsRRF in _olistItemsRRFCollection)
                                {
                                    sendEmailReleaseReject(ctx, _oweb, itemsRRF, EventId, To, Cc, TrackingId);
                                }

                            }
                            else if (items["Flag"].ToString() == "10")
                            {
                                SendResourceAllocationEmailWithOutRRF(ctx, _oweb, EventId, To, Cc, TrackingId);
                            }
                            else
                            {
                                List _olistRRFMaster = _oweb.Lists.GetByTitle("RRF");
                                var CamlQueryRRF = new CamlQuery() { ViewXml = "<View><Query><Where><Eq><FieldRef Name='RRFNO' /><Value Type='Text'>" + TrackingId + "</Value></Eq></Where></Query></View>" };
                                ListItemCollection _olistItemsRRFCollection = _olistRRFMaster.GetItems(CamlQueryRRF);
                                ctx.Load(_olistItemsRRFCollection);
                                ctx.ExecuteQuery();
                                foreach (ListItem itemsRRF in _olistItemsRRFCollection)
                                {
                                    if (items["Flag"].ToString() == "5")
                                    {
                                        SendEmail(ctx, _oweb, itemsRRF, EventId, To, Cc);
                                    }

                                    if (items["Flag"].ToString() == "6")
                                    {
                                        SendDynamicTableEmail(ctx, _oweb, itemsRRF, EventId, To, Cc);
                                    }
                                }
                            }
                            if (ErrorFlag != "1")
                            {
                                items["Flag"] = "0";
                            }


                            else
                            {
                                items["Flag"] = "20";

                                SendMailsDatBaseProfile("pankaj.singh@e2eprojects.com", "uday.s@e2eprojects.com", "Error AAya", "Error AAYA");
                            }
                            items.Update();
                            ctx.Load(items);
                            ctx.ExecuteQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SendMailsDatBaseProfile("pankaj.singh@e2eprojects.com", "uday.s@e2eprojects.com", ex.ToString(), "Error AAYA");
                Console.WriteLine(ex.Message);
            }
        }

        public static void SendResourceAllocationEmailWithOutRRF(ClientContext ctx, Web _oweb, string EventId, string FinalTo, string FinalCc, string TrackingId)
        {
            JArray jarr = null;
            MsOnlineClaimsHelper claimsHelper = new MsOnlineClaimsHelper(URL, UserName, Password);
            var AllocationStartDate = "";
            var AllocationEndDate = "";
            var AllocationPercentage = "";
            var AllocationBillableStatus = "";
            var AllocationProjectName = "";
            var AllocationCustomer = "";
            var AllocationProjectManager = "";
            var AllocationClientPartner = "";
            var AllocationProjectLocation = "";
            var AllocationAssociateContact = "";
            var AllocationProjectCode = "";
            string FinalToEmailId = "";
            string FinalCcEmailId = "";
            string finalsubject = "";
            string textBody = "";
            var ResourceEmpId = "";
            var AllocatedResourceName = "";
            try
            {


                var AllocatedProjectDetails = GetAllocatedResourceDetails(ctx, _oweb, TrackingId);
            



                if (AllocatedProjectDetails.Count() > 0)
                {
                    AllocationProjectCode = AllocatedProjectDetails[0]["AllocatedProjectCode"].ToString();
                    //int InternalProjectCode = Convert.ToInt32(AllocatedProjectDetails[0]["InternalID"]);
                    var AllocatedProjectCodeDetails = GetAllocatedProjectDetails(ctx, _oweb, AllocationProjectCode);
                    ResourceEmpId = AllocatedProjectDetails[0]["EmployeeID"].ToString();
                    AllocatedResourceName = AllocatedProjectDetails[0]["ResourceFullName"].ToString();


                    var AllocatedEmployeeStatus = GetAllocatedResourceStatus(ctx, _oweb, ResourceEmpId);

                    if (AllocatedProjectCodeDetails.Count() > 0)
                    {

                        AllocationStartDate = AllocatedProjectDetails[0]["Startdatetime"].ToString();

                        AllocationStartDate = Convert.ToDateTime(AllocationStartDate).ToShortDateString();

                        AllocationEndDate = AllocatedProjectDetails[0]["Finishdatetime"].ToString();

                        AllocationEndDate = Convert.ToDateTime(AllocationEndDate).ToShortDateString();

                        AllocationPercentage = AllocatedProjectDetails[0]["Allocation"].ToString();

                        if (AllocatedEmployeeStatus.Count() > 0)
                        {
                            AllocationBillableStatus = AllocatedEmployeeStatus[0]["EmployeeRole"].ToString();
                            //AllocationAssociateContact = AllocatedEmployeeStatus[0]["PhoneNumber"].ToString();
                            //if (AllocationAssociateContact == null)
                            //{
                            //    AllocationAssociateContact = "";
                            //}
                        }
                        AllocationProjectName = AllocatedProjectDetails[0]["ProjectName"].ToString();
                        AllocationCustomer = AllocatedProjectCodeDetails[0]["CustomerName"].ToString();
                        AllocationProjectManager = AllocatedProjectCodeDetails[0]["ProjectOwnerName"].ToString();

                        if(AllocationProjectManager!="")
                        {
                            var request = (HttpWebRequest)WebRequest.Create(URL + "_api/ProjectData/Resources?$select=ResourceName,PhoneNumber,ResourceEmailAddress&$filter=ResourceName eq '" + AllocationProjectManager + "'");
                            request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                            request.Method = WebRequestMethods.Http.Get;
                            request.Accept = "application/json;odata=verbose";
                            // request.ContentType = "application/json;odata=verbose";
                            request.ContentLength = 0;

                            var securePassword = new SecureString();
                            foreach (char c in Password)
                            {
                                securePassword.AppendChar(c);
                            }
                            request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);
                            /* HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                             endpointRequest.Method = "GET";
                             //if (XML == false)
                             endpointRequest.Accept = "application/json;odata=verbose";
                             endpointRequest.UseDefaultCredentials = false;

                             endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                            HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                            WebResponse webResponse = request.GetResponse();
                            Stream webStream = webResponse.GetResponseStream();
                            StreamReader responseReader = new StreamReader(webStream);
                            string response = responseReader.ReadToEnd();
                            JObject jobj = JObject.Parse(response);
                            jarr = (JArray)jobj["d"]["results"];
                            JArray jarrPT = new JArray();
                            foreach (JObject j in jarr)
                            {
                                JObject jPT = new JObject();
                                string ResourceName = j["ResourceName"].ToString();
                                 AllocationAssociateContact = j["PhoneNumber"].ToString();
                                 AllocationProjectManager = j["ResourceEmailAddress"].ToString();


                            }
                        }




                        AllocationClientPartner = AllocatedProjectCodeDetails[0]["ClientPartner"].ToString();
                        if (AllocationClientPartner != "")
                        {
                            AllocationClientPartner = AllocatedProjectCodeDetails[0]["ClientPartner"].ToString().Split('|')[1];
                        }

                        AllocationProjectLocation = AllocatedProjectDetails[0]["ProjectLocation"].ToString();
                        if (AllocationProjectLocation == null)
                        {
                            AllocationProjectLocation = "";
                        }


                        textBody = "<span style='font-size:11pt;font-family:'calibri'>Hi " + AllocatedResourceName +"</span>,</br></br>" +
                             "<span style='font-size:11pt;font-family:'calibri''>You have been assigned to  new project, The details are given below.</span></br></br></br>" +

                 "<table class='MsoTableGrid' cellspacing='0' cellpadding='0' width='100%' border='1' style='border-width: medium; border-style: none; border-color: initial; width: 559.7pt; margin: auto auto auto -0.25pt;'>" +
    "<tr>" +
                            
    "<td valign='top' width='37' style='border-width: 1pt; border-style: solid; border-color: windowtext; width: 27.8pt; background: #2f5496; padding: 0in 5.4pt;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='calibri'><font size='2'>Customer Name</font></font></span></p>" +
    "</td>" +
    "<td valign='top' width='72' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 54.35pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='calibri'><font size='2'>Project Name</font></font></span></p>" +
    "</td>" +
    "<td valign='top' width='88' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 66.35pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='calibri'><font size='2'>Project Code</font></font></span></p>" +
    "</td>" +

    "<td valign='top' width='25' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 18.85pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='calibri'><font size='2'>Project Manager (PM) Email</font></font></span></p>" +
    "</td>" +

    "<td valign='top' width='25' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 18.85pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='calibri'><font size='2'>PM Contact No</font></font></span></p>" +
    "</td>" +

    "<td valign='top' width='61' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 45.75pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='calibri'><font size='2'>Client Partner</font></font></span></p>" +
    "</td>" +
    "<td valign='top' width='77' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 57.65pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='calibri'><font size='2'>Allocation Start Date</font></font></span></p>" +
    "</td>" +
    "<td valign='top' width='62' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 46.6pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='calibri'><font size='2'>Allocation End Date</font></font></span></p>" +
    "</td>" +
    "<td valign='top' width='73' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 32.35pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='calibri'><font size='2'>% Allocation</font></font></span></p>" +
    "</td>" +
    "<td valign='top' width='73' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 32.35pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='calibri'><font size='2'>Project Location</font></font></span></p>" +
    "</td>" +
    "<td valign='top' width='73' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 32.35pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='calibri'><font size='2'>Billability</font></font></span></p>" + "</td>";



                        textBody += "<tr><td valign='top' width='37'style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 27.8pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: 1pt solid windowtext; background-color: transparent;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #1a1a1a;'><span style='font-family:'calibri';font-size:11pt;'>" + AllocationCustomer + "</span><br/></span></p>" +
    "</td>" +
    "<td valign='top' width='72' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 54.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
    "<span style='color: #1a1a1a;'><span style='font-family:'calibri';font-size:11pt;'>" + AllocationProjectName + "​</span><br/></span></td>" +
    "<td valign='top' width='88' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 66.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
    "<span style='color: #1a1a1a;'><span style='font-family:'calibri';font-size:11pt;'>" + AllocationProjectCode + "</span><br/></span></td>" +

    "<td valign='top' width='25' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 18.85pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
    "<span style='color: #1a1a1a;'><span style='font-family:'calibri';font-size:11pt;'> " + AllocationProjectManager + "</span><br/></span></td>" +

       "<td valign='top' width='25' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 18.85pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
    "<span style='color: #1a1a1a;'><span span style='font-family:'calibri';font-size:11pt;'> " + AllocationAssociateContact + "</span><br/></span></td>" +

    "<td valign='top' width='61' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 45.75pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #1a1a1a;'><span style='font-family:'calibri';font-size:11pt;'> " + AllocationClientPartner + "</span><br/></span></p>" +
    "</td>" +
    "<td valign='top' width='77' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 57.65pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
    "<span style='color: #1a1a1a;font-family:'calibri';font-size:11pt;'>" + AllocationStartDate + "</span>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "</p>" +
    "</td>" +
    "<td valign='top' width='62' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 46.6pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
    "<span style='color: #1a1a1a;'><span style='font-family:'calibri''> " + AllocationEndDate + "</span><br/></span></td>" +
    "<td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
    "<span style='color: #1a1a1a;'><span style='font-family:'calibri';font-size:11pt;'> " + AllocationPercentage + "</span><br/></span></td>" +
    "<td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
    "<span style='color: #1a1a1a;'><span style='font-family:'calibri';font-size:11pt;'> " + AllocationProjectLocation + "</span><br/></span></td>" +
    "<td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
    "<span style='color: #1a1a1a;'><span style='font-family:'calibri';font-size:11pt;'> " + AllocationBillableStatus + "</span><br/></span></td></tr>";

   



                        textBody += "</table><br /><br /><br />" +
                            "<span style='color: #1a1a1a;font-size:11pt;font-family:'calibri';'><font size='3'><b><u>Please note the following point for Billable Project Deployment-</u></b></font></br></br>" +
                    "<table class='MsoTableGrid'  cellspacing='0' cellpadding='0' width='150%'  border='1' style='border-collapse: collapse; border-width: medium;border-style: double;  width: 600pt; margin: auto auto auto -0.25pt;'> " +
                    "<tbody ><tr><td><span style='padding:5px;line-height:25px'><font-size:11pt;font-family:''calibri'';><b>% Allocation​</b></font></span></td><td><span style='padding:5px;line-height:25px'><font size='3'>As per the new policy, 90% allocation on billable project will be considered as fully billable. Hence you will be allocated max. 90% on the “Billable” project(s). 10% allocation is reserved for “Other” project assignment like Training, Practice work (like pre-sales, conducting interviews etc.) on need basis. Distribute your actual work time on both accordingly..</font></span></td></tr>" +
                            "<tr><span style='padding:5px;line-height:25px'><td><font-size:11pt;font-family:Calibri,sans-serif;margin:0 0 12pt 0;><b>Discuss with PM</b></font></span></td><td><span style='padding:5px;line-height:25px'><font size='3'>- Expected deliverables while on the project.</br>" +
                                                                                    "- About location of work, reporting time, dress code etc.</br>" +
                                                                                    "- Travel requirement at client place, if any, for the project needs.</font></span></td></tr>" +
                            "<tr><td><span style='padding:5px;line-height:25px'><font-size:11pt;font-family:calibri'><b>Contact Details</b></font></span></td><td><span style='padding:5px;line-height:25px'><font size='3'>Share your contact details with PM. Update your contact information in Workday to reflect Outlook contact card. The change in contact will reflect in max. 24 hours.</font></span></td></tr>" +

                            "<tr><td><span style='padding:5px;line-height:25px'><font-size:11pt;font-family:'calibri'><b>Project Tasks</b></font></span></td><td><span style='padding:5px;line-height:25px'><font size='3'>PM will assign you tasks in PPM. These tasks will be available in timesheet. Please check with PM if you do not see tasks to fill timesheet. The procedure for checking task assignment is available on link Solace.becone.com => RMO => How To => Check Task Assignment or <a href='https://bristleconeonline.sharepoint.com/:w:/r/RMORevamp/_layouts/15/Doc.aspx?sourcedoc=%7B8B84E6D3-5D87-49BD-90D1-4B0E6A13F4B3%7D&file=Steps%20to%20Check%20Project%20Allocation.docx&action=default&mobileredirect=true&cid=3b3ee741-a96e-4db5-8f93-d0e127b2cf3a'>Click here</a></font></span></td></tr>" +
                            "<tr><td><span style='padding:5px;line-height:25px'><font-size:11pt;font-family:'calibri'><b>TimeSheets</b></font></span></td><td><span style='padding:5px;line-height:25px'><font size='3'>- Timesheet submission compliance in an important KPI. 98%+ timely submission is required to get 5 rating. Hence, submit your Timesheets on PPM in timely manner.</br>" +
                                                                                "- Check your timesheet submission compliance% on RMO Portal every month.</br>" +
                                                                                "- Follow-up with PM if your timesheet are not approved for previous week by every Tuesday.</br>" +
                                                                                "- The previous months timesheet are locked on 5th business day of current month. Your PM has rights to reopen approved timesheet for corrections, if any, prior to global lock. However, it will affect your timesheet submission compliance.</br>"+
                                                                                "- The link for how to submit timesheet is available on Solace.becone.com => RMO => How To => Create Timesheet in PPM or  <a href='https://bristleconeonline.sharepoint.com/:w:/r/RMORevamp/_layouts/15/Doc.aspx?sourcedoc=%7B09E42B36-5D47-4122-9969-2B2D84791AC0%7D&file=Steps%20for%20Creating%20Timesheet.docx&action=default&mobileredirect=true&cid=da479efc-8312-4f20-b862-b6671bb47f68'>Click here</a></font></span></td></tr>" +
                            "<tr><td><span style='padding:5px;line-height:25px'><font-size:11pt;font-family:'calibri'><b>Project ​Release</b></font></span></td><td><span style='padding:5px;line-height:25px'><font size='3'>An advance notification would be sent to you prior to release/extension from the project. Please approach your PM once you receive such alerts. Seek “Project End Feedback” from your PM on PPM</font></span></td></tr> " +
                            "<tr><td><span style='padding:5px;line-height:25px'><font-size:11pt;font-family:'calibri'><b>Profile Update</b></font></span></td><td><span style='padding:5px;line-height:25px'>" +
                            "<font-size:11pt;font-family:'calibri'>" +
                            "You are expected to update one page PPTx & MS-Word resume in the first week of every quarter in PPM. The video for how to update resume is available on Solace.becone.com => RMO => How To => Update Skill <a href='https://web.microsoftstream.com/video/e41ab267-059e-4133-9272-4eeb491d2205?list=trending&referrer=https:%2F%2Fsolace.bcone.com%2F&referrer=https:%2F%2Fsolace.bcone.com%2F'>Click here</a>or To access profile update page on PPM <a href='https://bristleconeonline.sharepoint.com/sites/pwa/SitePages/MyProfile.aspx'>click here</a>" +
                            "</font></span></td></tr></tbody> </table>" +

                    "</br></br>" +
                    "<span style='font-size:11pt;font-family:'calibri';>Reach out to <a href='mailto:rmo@bcone.com'>rmo@bcone.com</a>	for further clarifications. </span></br></br>" +
                    "<span><span style='color:#cc6600;font-size:11pt;font-family:'calibri';'>All the Best for your new assignment.</span></span></br></br>" +
                    "<span><b>Thanks & Regards, </b></span></br>" +
                    "<span>RMO Team </span></br></br>";


                        finalsubject = "You have been assigned to " + AllocationProjectName;
                        int k = 0;
                        if (FinalTo != "")
                        {
                            string[] newTo = FinalTo.Split(';');
                            foreach (string EmployeeID in newTo)
                            {
                                if (EmployeeID.ToString().Contains('@'))
                                {
                                    if (k == 0)
                                    {
                                        FinalToEmailId = EmployeeID;
                                    }
                                    else
                                    {
                                        FinalToEmailId = FinalToEmailId + ";" + EmployeeID + ";";
                                    }
                                    k++;
                                }
                                else
                                {
                                    var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                                    request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                                    request.Method = WebRequestMethods.Http.Get;
                                    request.Accept = "application/json;odata=verbose";
                                    // request.ContentType = "application/json;odata=verbose";
                                    request.ContentLength = 0;

                                    var securePassword = new SecureString();
                                    foreach (char c in Password)
                                    {
                                        securePassword.AppendChar(c);
                                    }
                                    request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);
                                    /* HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                                     endpointRequest.Method = "GET";
                                     //if (XML == false)
                                     endpointRequest.Accept = "application/json;odata=verbose";
                                     endpointRequest.UseDefaultCredentials = false;

                                     endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                                    HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                                    WebResponse webResponse = request.GetResponse();
                                    Stream webStream = webResponse.GetResponseStream();
                                    StreamReader responseReader = new StreamReader(webStream);
                                    string response = responseReader.ReadToEnd();
                                    JObject jobj = JObject.Parse(response);
                                    jarr = (JArray)jobj["d"]["results"];
                                    JArray jarrPT = new JArray();
                                    foreach (JObject j in jarr)
                                    {
                                        JObject jPT = new JObject();
                                        string emailId = j["Email"].ToString();
                                        if (k == 0)
                                        {
                                            FinalToEmailId = emailId;
                                        }
                                        else
                                        {
                                            FinalToEmailId = FinalToEmailId + ";" + emailId + ";";
                                        }
                                        k++;
                                    }
                                }
                            }
                        }

                        if (FinalCc != "")
                        {
                            string[] newCo = FinalCc.Split(';');
                            int l = 0;
                            foreach (string EmployeeID in newCo)
                            {
                                if (EmployeeID.ToString().Contains('@'))
                                {
                                    if (l == 0)
                                    {
                                        FinalCcEmailId = EmployeeID;
                                    }
                                    else
                                    {
                                        FinalCcEmailId = FinalCcEmailId + ";" + EmployeeID + ";";
                                    }
                                    l++;
                                }
                                else
                                {
                                    var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                                    request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                                    request.Method = WebRequestMethods.Http.Get;
                                    request.Accept = "application/json;odata=verbose";
                                    // request.ContentType = "application/json;odata=verbose";
                                    request.ContentLength = 0;

                                    var securePassword = new SecureString();
                                    foreach (char c in Password)
                                    {
                                        securePassword.AppendChar(c);
                                    }
                                    request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);

                                    /*  HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                                      endpointRequest.Method = "GET";
                                      //if (XML == false)
                                      endpointRequest.Accept = "application/json;odata=verbose";
                                      endpointRequest.UseDefaultCredentials = false;

                                      endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                                    HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                                    WebResponse webResponse = request.GetResponse();
                                    Stream webStream = webResponse.GetResponseStream();
                                    StreamReader responseReader = new StreamReader(webStream);
                                    string response = responseReader.ReadToEnd();
                                    JObject jobj = JObject.Parse(response);
                                    jarr = (JArray)jobj["d"]["results"];
                                    JArray jarrPT = new JArray();
                                    foreach (JObject j in jarr)
                                    {
                                        JObject jPT = new JObject();
                                        string emailId = j["Email"].ToString();
                                        if (l == 0)
                                        {
                                            FinalCcEmailId = emailId;
                                        }
                                        else
                                        {
                                            FinalCcEmailId = FinalCcEmailId + ";" + emailId + ";";
                                        }
                                        l++;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorFlag = "1";
                // throw;
                SendMailsDatBaseProfile("pankaj.singh@e2eprojects.com", "uday.s@e2eprojects.com", ex.ToString(), "Error SendResourceAllocationEmailWithOutRRF");
            }

            if (ErrorFlag != "1")
            {

                //SqlCommand cmdExec = new SqlCommand("exec msdb.dbo.sp_send_dbmail @Profile_name=@Profile_name1," +
                //                         "@recipients=@recipients1,@copy_recipients=@copy_recipients1,@subject=@subject1,@body=@body1,@body_format=@body_format1", con);
                //if (con.State == ConnectionState.Closed)
                //{
                //    con.Open();
                //}
                //try
                //{
                //    //FinalToEmailId = "uday.s@e2eprojects.com";
                //    //FinalCcEmailId = "pankaj.singh@e2eprojects.com";
                //    cmdExec.Parameters.AddWithValue("@Profile_name1", "RMO");
                //    cmdExec.Parameters.AddWithValue("@recipients1", FinalToEmailId);
                //    cmdExec.Parameters.AddWithValue("@subject1", finalsubject);
                //    cmdExec.Parameters.AddWithValue("@body1", textBody);
                //    //cmdExec.Parameters.AddWithValue("@blind_copy_recipients1", "uday.s@e2eprojects.com");
                //    cmdExec.Parameters.AddWithValue("@copy_recipients1", FinalCcEmailId);
                //    cmdExec.Parameters.AddWithValue("@body_format1", "HTML");
                //    cmdExec.ExecuteNonQuery();
                //}
                //catch (Exception ex)
                //{
                //    Console.WriteLine(ex.Message);
                //    SendMailsDatBaseProfile("pankaj.singh@e2eprojects.com", "uday.s@e2eprojects.com", ex.ToString(), "Error SendResourceAllocationEmailWithOutRRF");
                //}
                SendMailsDatBaseProfile(FinalToEmailId, FinalCcEmailId, textBody, finalsubject);

            }

        }

        public static ClientContext GetContext()
        {
           
            SecureString passWord = new SecureString();
            string PWAUrl = URL;
            Web web;
            ClientContext ctx = new ClientContext(PWAUrl);
            using (ctx)
            {
                try
                {
                    foreach (char c in Password.ToCharArray()) passWord.AppendChar(c);
                    ctx.Credentials = new SharePointOnlineCredentials(UserName, passWord);
                    web = ctx.Web;
                    ctx.Load(web);
                    ctx.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    Console.Write(ex.Message);
                    // throw;
                }

            }
            return ctx;
        }
        public static JToken GetAllocatedResourceDetails(ClientContext ctx, Web _oweb, string TrackingId)
        {
            using (var client = new WebClient())
            {
                int InternalId = Convert.ToInt32(TrackingId);
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
                var endpointUri = new Uri(serviceURL + "ProjectWiseResourceAllocation?$filter=InternalID eq " + InternalId + "&$orderby=created_date desc&$top=1");
                var result = client.DownloadString(endpointUri);
                var t = JToken.Parse(result);
                return t["d"]["results"];
            }
        }
        

        public static JToken GetAllocatedProjectDetails(ClientContext ctx, Web _oweb, string TrackingId)
        {
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
                var endpointUri = new Uri(serviceURL + "Projects?$filter=ProjectCode eq '" + TrackingId + "'&$top=1");
                var result = client.DownloadString(endpointUri);
                var t = JToken.Parse(result);
                return t["d"]["results"];
            }
           

        }


        public static void SendDynamicTableEmail(ClientContext ctx, Web _oweb, ListItem itemsRRF, string EventId, string FinalTo, string FinalCc)
        {
            JArray jarr = null;
            MsOnlineClaimsHelper claimsHelper = new MsOnlineClaimsHelper(URL, UserName, Password);
            List _olist = _oweb.Lists.GetByTitle("BconeEmailConfiguration");
            CamlQuery camlqueryConfig = new CamlQuery();
            camlqueryConfig.ViewXml = "<View><Query><Where><Eq><FieldRef Name='EventID' /><Value Type='Text'>" + EventId + "</Value></Eq></Where></Query></View>";
            ListItemCollection EmailConfigurationtItemsCollection = _olist.GetItems(camlqueryConfig);
            ctx.Load(EmailConfigurationtItemsCollection);
            ctx.ExecuteQuery();
            string FinalToEmailId = "";
            int k = 0;
            if (FinalTo != "")
            {
                string[] newTo = FinalTo.Split(';');
                foreach (string EmployeeID in newTo)
                {
                    if (EmployeeID.ToString().Contains('@'))
                    {
                        if (k == 0)
                        {
                            FinalToEmailId = EmployeeID;
                        }
                        else
                        {
                            FinalToEmailId = FinalToEmailId + ";" + EmployeeID + ";";
                        }
                        k++;
                    }
                    else
                    {
                        var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                        request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                        request.Method = WebRequestMethods.Http.Get;
                        request.Accept = "application/json;odata=verbose";
                        // request.ContentType = "application/json;odata=verbose";
                        request.ContentLength = 0;

                        var securePassword = new SecureString();
                        foreach (char c in Password)
                        {
                            securePassword.AppendChar(c);
                        }
                        request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);
                        /* HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                         endpointRequest.Method = "GET";
                         //if (XML == false)
                         endpointRequest.Accept = "application/json;odata=verbose";
                         endpointRequest.UseDefaultCredentials = false;

                         endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                        HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                        WebResponse webResponse = request.GetResponse();
                        Stream webStream = webResponse.GetResponseStream();
                        StreamReader responseReader = new StreamReader(webStream);
                        string response = responseReader.ReadToEnd();
                        JObject jobj = JObject.Parse(response);
                        jarr = (JArray)jobj["d"]["results"];
                        JArray jarrPT = new JArray();
                        foreach (JObject j in jarr)
                        {
                            JObject jPT = new JObject();
                            string emailId = j["Email"].ToString();
                            if (k == 0)
                            {
                                FinalToEmailId = emailId;
                            }
                            else
                            {
                                FinalToEmailId = FinalToEmailId + ";" + emailId + ";";
                            }
                            k++;
                        }
                    }
                }
            }
            string FinalCcEmailId = "";
            if (FinalCc != "")
            {
                string[] newCo = FinalCc.Split(';');
                int l = 0;
                foreach (string EmployeeID in newCo)
                {
                    if (EmployeeID.ToString().Contains('@'))
                    {
                        if (l == 0)
                        {
                            FinalCcEmailId = EmployeeID;
                        }
                        else
                        {
                            FinalCcEmailId = FinalCcEmailId + ";" + EmployeeID + ";";
                        }
                        l++;
                    }
                    else
                    {
                        var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                        request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                        request.Method = WebRequestMethods.Http.Get;
                        request.Accept = "application/json;odata=verbose";
                        // request.ContentType = "application/json;odata=verbose";
                        request.ContentLength = 0;

                        var securePassword = new SecureString();
                        foreach (char c in Password)
                        {
                            securePassword.AppendChar(c);
                        }
                        request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);
                        /*HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                        endpointRequest.Method = "GET";
                        //if (XML == false)
                        endpointRequest.Accept = "application/json;odata=verbose";
                        endpointRequest.UseDefaultCredentials = false;

                        endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                        HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                        WebResponse webResponse = request.GetResponse();
                        Stream webStream = webResponse.GetResponseStream();
                        StreamReader responseReader = new StreamReader(webStream);
                        string response = responseReader.ReadToEnd();
                        JObject jobj = JObject.Parse(response);
                        jarr = (JArray)jobj["d"]["results"];
                        JArray jarrPT = new JArray();
                        foreach (JObject j in jarr)
                        {
                            JObject jPT = new JObject();
                            string emailId = j["Email"].ToString();
                            if (l == 0)
                            {
                                FinalCcEmailId = emailId;
                            }
                            else
                            {
                                FinalCcEmailId = FinalCcEmailId + ";" + emailId + ";";
                            }
                            l++;
                        }
                    }
                }
            }

            string body = Convert.ToString(EmailConfigurationtItemsCollection[0]["Body"]);
            string Subject = Convert.ToString(EmailConfigurationtItemsCollection[0]["Subject"]);
            string noHTMLBody = System.Text.RegularExpressions.Regex.Replace(body, @"<[^>]+>|&nbsp;", "").Trim();
            string noHTMLNormalisedBody = System.Text.RegularExpressions.Regex.Replace(noHTMLBody, @"\s{2,}", " ");
            string noHTMLSubject = System.Text.RegularExpressions.Regex.Replace(Subject, @"<[^>]+>|&nbsp;", "").Trim();
            string noHTMLNormalisednoHTMLSubject = System.Text.RegularExpressions.Regex.Replace(noHTMLSubject, @"\s{2,}", " ");
            StringBuilder stringBuilder = new StringBuilder(body);
            StringBuilder stringBuildersubject = new StringBuilder(Subject);

            List<string> Bodyvariable = ExtractFromString(noHTMLNormalisedBody, "&#123;", "&#125;");

            ReplaceVariablevalue(itemsRRF, ctx, _oweb, ref _olist, ref camlqueryConfig, stringBuilder, Bodyvariable, "Body", FinalTo, EventId);

            List<string> Subjectvariable = ExtractFromString(noHTMLNormalisednoHTMLSubject, "{", "}");

            ReplaceVariablevalue(itemsRRF, ctx, _oweb, ref _olist, ref camlqueryConfig, stringBuildersubject, Subjectvariable, "Subject", FinalTo, EventId);

            List<string> Hyperlink = ExtractFromString(noHTMLNormalisedBody, "&lt;", "&gt;");

            for (int i = 0; i < Hyperlink.Count; i++)
            {
                string variablename = Hyperlink[i];
                stringBuilder.Replace("&lt;" + variablename + "&gt;", "<a style='color:#000000;face:Segoe UI Semibold, Calibri' href='" + RedirectURL + "'>Click Here </a>");

            }

            string Body = string.Empty;

            Body = stringBuilder.ToString();

            string finalsubject = stringBuildersubject.ToString();

            string textBody = "";
            List _olistResourceAll = _oweb.Lists.GetByTitle("RMOResourceAllocation");
            CamlQuery camlquery = new CamlQuery();
            camlquery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='RRFNumber' /><Value Type='Text'>" + itemsRRF["RRFNO"] + "</Value></Eq></Where></Query></View>";
            ListItemCollection ResourceAllConfigurationtItemsCollection = _olistResourceAll.GetItems(camlquery);
            ctx.Load(ResourceAllConfigurationtItemsCollection);
            ctx.ExecuteQuery();
            if (ResourceAllConfigurationtItemsCollection.Count > 0)
            {
                string EmployeeCode = "";
                string semicolon = ";";
                if (ResourceAllConfigurationtItemsCollection[0]["AllocatedResource"] != null)
                {
                    EmployeeCode = Convert.ToString(ResourceAllConfigurationtItemsCollection[0]["AllocatedResource"]);
                    if (EmployeeCode.IndexOf('\t') > -1)
                    {
                        EmployeeCode = EmployeeCode.Replace("\t", "");
                    }
                }
                else if (ResourceAllConfigurationtItemsCollection[0]["ShortlistedResource"] != null)
                {
                    EmployeeCode = Convert.ToString(ResourceAllConfigurationtItemsCollection[0]["ShortlistedResource"]);
                    if (EmployeeCode.IndexOf('\t') > -1)
                    {
                        EmployeeCode = EmployeeCode.Replace("\t", "");
                    }
                }
                else if (ResourceAllConfigurationtItemsCollection[0]["suggestedResource"] != null)
                {
                    EmployeeCode = Convert.ToString(ResourceAllConfigurationtItemsCollection[0]["suggestedResource"]);
                    if (EmployeeCode.IndexOf('\t') > -1)
                    {
                        EmployeeCode = EmployeeCode.Replace("\t", "");
                    }
                }
                textBody = "" + Body + "<br />" +
                "<table class='MsoTableGrid' cellspacing='0' cellpadding='0' width='100%' border='1' style='border-width: medium; border-style: none; border-color: initial; width: 559.7pt; margin: auto auto auto -0.25pt;'>" +
"<tr>" +
"<td valign='top' width='37' style='border-width: 1pt; border-style: solid; border-color: windowtext; width: 27.8pt; background: #2f5496; padding: 0in 5.4pt;'>" +
"<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
"<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Sr No</font></font></span></p>" +
"</td>" +
"<td valign='top' width='72' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 54.35pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
"<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
"<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>RRF Number</font></font></span></p>" +
"</td>" +
"<td valign='top' width='88' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 66.35pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
"<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
"<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Raised On</font></font></span></p>" +
"</td>" +
"<td valign='top' width='25' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 18.85pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
"<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
"<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Proposed Emp ID</font></font></span></p>" +
"</td>" +
"<td valign='top' width='61' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 45.75pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
"<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
"<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Proposed Employee</font></font></span></p>" +
"</td>" +
"<td valign='top' width='77' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 57.65pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
"<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
"<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Proposed Customer</font></font></span></p>" +
"</td>" +
"<td valign='top' width='62' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 46.6pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
"<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
"<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Proposed Project</font></font></span></p>" +
"</td>" +
"<td valign='top' width='73' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 32.35pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
"<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
"<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Soft Blocked Date</font></font></span></p>" +
"</td>";
                int o = 1;
                string rrfno = "";
                string createdDate = "";
                string EmployeeId = "";
                string ResourceFullName = "";
                string Customer = "";
                string ProjectName = "";
                string Created_On = "";
                if (EmployeeCode.IndexOf('\t') > -1)
                {
                    EmployeeCode = EmployeeCode.Replace("\t", "");
                }

                if (EmployeeCode.Contains(semicolon))
                {
                    string[] EmpIDValue = EmployeeCode.Split(';');
                    foreach (var empid in EmpIDValue)
                    {

                        var listitem = GetList(ctx, _oweb, itemsRRF, empid);
                        var listitemSoftBlock = GetListSoftBlock(ctx, _oweb, itemsRRF, empid);
                        if (listitem.Count() > 0)
                        {
                            if (itemsRRF["RRFNO"] != null)
                                rrfno = Convert.ToString(itemsRRF["RRFNO"]);
                            if (EventId == "50")
                            {
                                if (itemsRRF["SubmittedDate"] != null && itemsRRF["SubmittedDate"].ToString() != "")
                                    createdDate = Convert.ToDateTime(itemsRRF["SubmittedDate"]).ToShortDateString();
                            }
                            else
                            {

                                if (itemsRRF["Created"] != null && itemsRRF["Created"].ToString() != "")
                                    createdDate = Convert.ToDateTime(itemsRRF["Created"]).ToShortDateString();
                            }

                            if (listitem[0]["EmployeeID"] != null)
                                EmployeeId = Convert.ToString(listitem[0]["EmployeeID"]);
                            if (listitem[0]["ResourceName"] != null)
                                ResourceFullName = Convert.ToString(listitem[0]["ResourceName"]);
                            if (itemsRRF["Customer"] != null)
                                Customer = Convert.ToString(itemsRRF["Customer"]);
                            if (itemsRRF["ProjectName"] != null)
                                ProjectName = Convert.ToString(itemsRRF["ProjectName"]);

                            if (listitemSoftBlock.Count() > 0)
                            {
                                if (listitemSoftBlock[0]["created_date"] != null && listitemSoftBlock[0]["created_date"].ToString() != "null" && listitemSoftBlock[0]["created_date"].ToString() != "")
                                    Created_On = Convert.ToDateTime(listitemSoftBlock[0]["created_date"]).ToShortDateString();
                            }
                            else
                            {
                                Created_On = "";
                            }
                            textBody += "<tr><td valign='top' width='37' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 27.8pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: 1pt solid windowtext; background-color: transparent;'>" +
"<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
"<span style='color: #1a1a1a;'><span>" + o + "</span><br/></span></p>" +
"</td>" +
"<td valign='top' width='72' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 54.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
"<span style='color: #1a1a1a;'><span>" + rrfno + "​</span><br/></span></td>" +
"<td valign='top' width='88' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 66.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
"<span style='color: #1a1a1a;'><span>" + createdDate + "</span><br/></span></td>" +
"<td valign='top' width='25' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 18.85pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
"<span style='color: #1a1a1a;'><span> " + EmployeeId + "</span><br/></span></td>" +
"<td valign='top' width='61' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 45.75pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
"<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
"<span style='color: #1a1a1a;'><span> " + ResourceFullName + "</span><br/></span></p>" +
"</td>" +
"<td valign='top' width='77' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 57.65pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
"<span style='color: #1a1a1a;'>" + Customer + "</span>" +
"<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
"</p>" +
"</td>" +
"<td valign='top' width='62' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 46.6pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
"<span style='color: #1a1a1a;'><span> " + ProjectName + "</span><br/></span></td>" +
"<td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
"<span style='color: #1a1a1a;'><span> " + Created_On + "</span><br/></span></td></tr>";
                            o++;
                        }
                    }
                }
                else
                {
                    var listitem = GetList(ctx, _oweb, itemsRRF, EmployeeCode);
                    var listitemSoftBlock = GetListSoftBlock(ctx, _oweb, itemsRRF, EmployeeCode);
                    if (listitem.Count() > 0)
                    {
                        if (itemsRRF["RRFNO"] != null)
                            rrfno = Convert.ToString(itemsRRF["RRFNO"]);
                        if (EventId == "50")
                        {
                            if (itemsRRF["SubmittedDate"] != null)
                                createdDate = Convert.ToDateTime(itemsRRF["SubmittedDate"]).ToShortDateString();
                        }
                        else if (EventId == "51")
                        {
                            if (itemsRRF["RRFCreatedDate"] != null)
                                createdDate = Convert.ToDateTime(itemsRRF["RRFCreatedDate"]).ToString("yyyy-MM-dd");


                        }
                        else
                        {
                            if (itemsRRF["Created"] != null)
                                createdDate = Convert.ToDateTime(itemsRRF["Created"]).ToShortDateString();
                        }

                        if (listitem[0]["EmployeeID"] != null)
                            EmployeeId = Convert.ToString(listitem[0]["EmployeeID"]);
                        if (listitem[0]["ResourceName"] != null)
                            ResourceFullName = Convert.ToString(listitem[0]["ResourceName"]);
                        if (itemsRRF["Customer"] != null)
                            Customer = Convert.ToString(itemsRRF["Customer"]);
                        if (itemsRRF["ProjectName"] != null)
                            ProjectName = Convert.ToString(itemsRRF["ProjectName"]);

                        if (listitemSoftBlock.Count() > 0)
                        {
                            if (listitemSoftBlock[0]["created_date"] != null && listitemSoftBlock[0]["created_date"].ToString() != "null" && listitemSoftBlock[0]["created_date"].ToString() != "")
                                Created_On = Convert.ToDateTime(listitemSoftBlock[0]["created_date"]).ToShortDateString();
                        }
                        else
                        {
                            Created_On = "";
                        }
                        textBody += "<tr><td>" + o + "</td><td> " + rrfno + "</td><td> " + createdDate + "</td><td> " + EmployeeId + "</td><td> " + ResourceFullName + "</td><td> " + Customer + "</td><td> " + ProjectName + "</td><td> " + Created_On + "</td></tr>";

                    }
                }
                textBody += "</table><br />" +
"</tr><br/><font size='2'>" +
         "<font color='#cc6600'>This message is auto-generated and do not reply to this email.​<br/><br/></font><strong><font color='#000000'>Thanks &amp; Regards,</font></strong><font color='#000000'></font><br/>RMO Team</font></p><br /><br />";
            }


            try
            {
                //SqlCommand cmdExec = new SqlCommand("exec msdb.dbo.sp_send_dbmail @Profile_name=@Profile_name1," +
                //                                "@recipients=@recipients1,@copy_recipients=@copy_recipients1,@subject=@subject1,@body=@body1,@body_format=@body_format1", con);
                //if (con.State == ConnectionState.Closed)
                //{
                //    con.Open();
                //}
                ////FinalToEmailId = "uday.s@e2eprojects.com";
                ////FinalCcEmailId = "pankaj.singh@e2eprojects.com";
                //cmdExec.Parameters.AddWithValue("@Profile_name1", "RMO");
                //cmdExec.Parameters.AddWithValue("@recipients1", FinalToEmailId);
                //cmdExec.Parameters.AddWithValue("@subject1", finalsubject);
                //cmdExec.Parameters.AddWithValue("@body1", textBody);
                //cmdExec.Parameters.AddWithValue("@copy_recipients1", FinalCcEmailId);
                ////cmdExec.Parameters.AddWithValue("@blind_copy_recipients1", "uday.s@e2eprojects.com");

                //cmdExec.Parameters.AddWithValue("@body_format1", "HTML");
                //cmdExec.ExecuteNonQuery();

                SendMailsDatBaseProfile(FinalToEmailId, FinalCcEmailId, textBody, finalsubject);

            }
            catch (Exception ex)
            {
                SendMailsDatBaseProfile("pankaj.singh@e2eprojects.com", "uday.s@e2eprojects.com", ex.ToString(), "Error SendDynamicTableEmail");
                // throw;
            }


        }


        public static void SendEmail(ClientContext ctx, Web _oweb, ListItem itemsRRF, string EventId, string FinalTo, string FinalCc)
        {
            try
            {
                JArray jarr = null;
                MsOnlineClaimsHelper claimsHelper = new MsOnlineClaimsHelper(URL, UserName, Password);
                List _olist = _oweb.Lists.GetByTitle("BconeEmailConfiguration");
                CamlQuery camlquery = new CamlQuery();
                camlquery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='EventID' /><Value Type='Text'>" + EventId + "</Value></Eq></Where></Query></View>";
                ListItemCollection EmailConfigurationtItemsCollection = _olist.GetItems(camlquery);
                ctx.Load(EmailConfigurationtItemsCollection);
                ctx.ExecuteQuery();
                string FinalToEmailId = "";
                int m = 0;
                if (FinalTo != "")
                {
                    string[] newTo = FinalTo.Split(';');
                    foreach (string EmployeeID in newTo)
                    {
                        if (EmployeeID.ToString().Contains('@'))
                        {
                            if (m == 0)
                            {
                                FinalToEmailId = EmployeeID;
                            }
                            else
                            {
                                FinalToEmailId = FinalToEmailId + ";" + EmployeeID + ";";
                            }
                            m++;
                        }
                        else
                        {
                            var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                            request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                            request.Method = WebRequestMethods.Http.Get;
                            request.Accept = "application/json;odata=verbose";
                            // request.ContentType = "application/json;odata=verbose";
                            request.ContentLength = 0;

                            var securePassword = new SecureString();
                            foreach (char c in Password)
                            {
                                securePassword.AppendChar(c);
                            }
                            request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);
                            /* HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                             endpointRequest.Method = "GET";
                             //if (XML == false)
                             endpointRequest.Accept = "application/json;odata=verbose";
                             endpointRequest.UseDefaultCredentials = false;

                             endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                            HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                            WebResponse webResponse = request.GetResponse();
                            Stream webStream = webResponse.GetResponseStream();
                            StreamReader responseReader = new StreamReader(webStream);
                            string response = responseReader.ReadToEnd();
                            JObject jobj = JObject.Parse(response);
                            jarr = (JArray)jobj["d"]["results"];
                            JArray jarrPT = new JArray();
                            foreach (JObject j in jarr)
                            {
                                JObject jPT = new JObject();
                                string emailId = j["Email"].ToString();
                                if (m == 0)
                                {
                                    FinalToEmailId = emailId;
                                    ResourceEmailValue = emailId;
                                }
                                else
                                {
                                    FinalToEmailId = FinalToEmailId + ";" + emailId + ";";
                                }
                                m++;
                            }
                        }
                    }
                }
                string FinalCcEmailId = "";
                if (FinalCc != "")
                {
                    string[] newCo = FinalCc.Split(';');
                    int o = 0;
                    foreach (string EmployeeID in newCo)
                    {
                        if (EmployeeID.ToString().Contains('@'))
                        {
                            if (o == 0)
                            {
                                FinalCcEmailId = EmployeeID;
                            }
                            else
                            {
                                FinalCcEmailId = FinalCcEmailId + ";" + EmployeeID + ";";
                            }
                            o++;
                        }
                        else
                        {
                            var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                            request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                            request.Method = WebRequestMethods.Http.Get;
                            request.Accept = "application/json;odata=verbose";
                            // request.ContentType = "application/json;odata=verbose";
                            request.ContentLength = 0;

                            var securePassword = new SecureString();
                            foreach (char c in Password)
                            {
                                securePassword.AppendChar(c);
                            }
                            request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);

                            /* HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                             endpointRequest.Method = "GET";
                             //if (XML == false)
                             endpointRequest.Accept = "application/json;odata=verbose";
                             endpointRequest.UseDefaultCredentials = false;

                             endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                            HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                            WebResponse webResponse = request.GetResponse();
                            Stream webStream = webResponse.GetResponseStream();
                            StreamReader responseReader = new StreamReader(webStream);
                            string response = responseReader.ReadToEnd();
                            JObject jobj = JObject.Parse(response);
                            jarr = (JArray)jobj["d"]["results"];
                            JArray jarrPT = new JArray();
                            foreach (JObject j in jarr)
                            {
                                JObject jPT = new JObject();
                                string emailId = j["Email"].ToString();
                                if (o == 0)
                                {
                                    FinalCcEmailId = emailId;
                                }
                                else
                                {
                                    FinalCcEmailId = FinalCcEmailId + ";" + emailId + ";";
                                }
                                o++;
                            }
                        }
                    }
                }
                string body = Convert.ToString(EmailConfigurationtItemsCollection[0]["Body"]);
                var pro = "https://bristleconeonline.sharepoint.com/:w:/r/RMORevamp/_layouts/15/Doc.aspx?sourcedoc=%7B8B84E6D3-5D87-49BD-90D1-4B0E6A13F4B3%7D&file=Steps%20to%20Check%20Project%20Allocation.docx&action=default&mobileredirect=true&cid=3b3ee741-a96e-4db5-8f93-d0e127b2cf3a";
                var my = "https://bristleconeonline.sharepoint.com/sites/pwa/SitePages/MyProfile.aspx#";
                var myone = "https://bristleconeonline.sharepoint.com/:w:/r/RMORevamp/_layouts/15/Doc.aspx?sourcedoc=%7B09E42B36-5D47-4122-9969-2B2D84791AC0%7D&file=Steps%20for%20Creating%20Timesheet.docx&action=default&mobileredirect=true&cid=da479efc-8312-4f20-b862-b6671bb47f68";
                var mytwo = "https://web.microsoftstream.com/video/e41ab267-059e-4133-9272-4eeb491d2205?list=trending&referrer=https:%2F%2Fsolace.bcone.com%2F&referrer=https:%2F%2Fsolace.bcone.com%2F";

                body = body.Replace("projecttask", " <a href='" + pro + "'>Click here </a>");
                body = body.Replace("myclick", "<a href='" + mytwo + "'>Click here </a>");
                body = body.Replace("Myprofile", "<a href='" + my + "'>Click here </a>");
                body = body.Replace("Mytimesheet", "<a href='" + myone + "'>Click here </a>");


                
               



                string Subject = Convert.ToString(EmailConfigurationtItemsCollection[0]["Subject"]);
                string noHTMLBody = System.Text.RegularExpressions.Regex.Replace(body, @"<[^>]+>|&nbsp;", "").Trim();
                string noHTMLNormalisedBody = System.Text.RegularExpressions.Regex.Replace(noHTMLBody, @"\s{2,}", " ");

                
                string noHTMLSubject = System.Text.RegularExpressions.Regex.Replace(Subject, @"<[^>]+>|&nbsp;", "").Trim();
                string noHTMLNormalisednoHTMLSubject = System.Text.RegularExpressions.Regex.Replace(noHTMLSubject, @"\s{2,}", " ");
                StringBuilder stringBuilder = new StringBuilder(body);
                StringBuilder stringBuildersubject = new StringBuilder(Subject);

                List<string> Bodyvariable = ExtractFromString(noHTMLNormalisedBody, "&#123;", "&#125;");

                ReplaceVariablevalue(itemsRRF, ctx, _oweb, ref _olist, ref camlquery, stringBuilder, Bodyvariable, "Body", FinalTo, EventId);

                List<string> Subjectvariable = ExtractFromString(noHTMLNormalisednoHTMLSubject, "{", "}");

                ReplaceVariablevalue(itemsRRF, ctx, _oweb, ref _olist, ref camlquery, stringBuildersubject, Subjectvariable, "Subject", FinalTo, EventId);

                List<string> Hyperlink = ExtractFromString(noHTMLNormalisedBody, "&lt;", "&gt;");

                for (int i = 0; i < Hyperlink.Count; i++)
                {
                    string variablename = Hyperlink[i];
                    stringBuilder.Replace("&lt;" + variablename + "&gt;", "<a style='color:#000000;fac" +
                        "e:Segoe UI Semibold, Calibri' href='" + RedirectURL + "'>Click Here </a>");

                }

                string Body = string.Empty;

                Body = stringBuilder.ToString();

                string finalsubject = stringBuildersubject.ToString();

                //SqlCommand cmdExec = new SqlCommand("exec msdb.dbo.sp_send_dbmail @Profile_name=@Profile_name1," +
                //                 "@recipients=@recipients1,@copy_recipients=@copy_recipients1,@subject=@subject1,@body=@body1,@body_format=@body_format1", con);
                //if (con.State == ConnectionState.Closed)
                //{
                //    con.Open();
                //}
                //FinalToEmailId = "pankaj.singh@e2eprojects.com";
                //FinalCcEmailId = "uday.s@e2eprojects.com";
                //cmdExec.Parameters.AddWithValue("@Profile_name1", "RMO");
                //cmdExec.Parameters.AddWithValue("@recipients1", FinalToEmailId);
                //cmdExec.Parameters.AddWithValue("@subject1", finalsubject);
                //cmdExec.Parameters.AddWithValue("@body1", Body);
                //cmdExec.Parameters.AddWithValue("@copy_recipients1", FinalCcEmailId);
                ////cmdExec.Parameters.AddWithValue("@blind_copy_recipients1", "Shweta.M@e2eprojects.com");
                //cmdExec.Parameters.AddWithValue("@body_format1", "HTML");
                //cmdExec.ExecuteNonQuery();

                SendMailsDatBaseProfile(FinalToEmailId, FinalCcEmailId, Body, finalsubject);
            }
            catch (Exception ex)
            {
                SendMailsDatBaseProfile("pankaj.singh@e2eprojects.com", "uday.s@e2eprojects.com", ex.ToString(), "Error SendEmail");
                Console.WriteLine(ex.Message);
            }
        }


        public static JToken GetList(ClientContext ctx, Web _oweb, ListItem itemsRRF, string EmpID)
        {

            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
                var endpointUri = new Uri(serviceURL + "Resources?$filter=EmployeeID eq '" + EmpID + "'and EmployeeStatus  ne 'Terminated'&$top=1");
                var result = client.DownloadString(endpointUri);
                var t = JToken.Parse(result);
                return t["d"]["results"];
            }


        }



        public static JToken GetListSoftBlock(ClientContext ctx, Web _oweb, ListItem itemsRRF, string EmpID)
        {
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
                var endpointUri = new Uri(serviceURL + "ProjectWiseResourceAllocation?$filter=EmployeeID eq '" + EmpID + "' and AllocatedProjectCode eq '" + itemsRRF["ProjectCode"] + "' and Flag eq 'Soft Block'&$orderby=created_date desc&$top=1");
                var result = client.DownloadString(endpointUri);
                var t = JToken.Parse(result);
                return t["d"]["results"];
            }
        }


        private static List<string> ExtractFromString(string text, string startString, string endString)
        {
            List<string> matched = new List<string>();
            int indexStart = 0, indexEnd = 0;
            bool exit = false;
            while (!exit)
            {
                indexStart = text.IndexOf(startString);
                indexEnd = text.IndexOf(endString);
                if (indexStart != -1 && indexEnd != -1)
                {
                    try {
                        matched.Add(text.Substring(indexStart + startString.Length,
                            indexEnd - indexStart - startString.Length));
                        text = text.Substring(indexEnd + endString.Length);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }

                }
                else
                    exit = true;
            }
            return matched;
        }


        public static JToken GetAllocatedResourceStatus(ClientContext ctx, Web _oweb, string EmpID)
        {
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
                var endpointUri = new Uri(serviceURL + "Resources?$filter=EmployeeID eq '" + EmpID + "'and EmployeeStatus ne 'Terminated'&$top=1");
                var result = client.DownloadString(endpointUri);
                var t = JToken.Parse(result);
                return t["d"]["results"];
            }
        }

        public static JToken GetAllocatedResourceName(ClientContext ctx, Web _oweb, string EmpID)
        {
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
                var endpointUri = new Uri(serviceURL + "Resources??$select=RoleBand,ResourceName&$filter=EmployeeID eq '" + EmpID + "'");
                var result = client.DownloadString(endpointUri);
                var t = JToken.Parse(result);
                return t["d"]["results"];
            }
        }






        public static string GetBody_EarlyRelaseReject(string Id_traking, ClientContext ctx, Web _oweb, ref List _olist, ref CamlQuery camlquery)
        {
            string bodyreturn = "";
            MsOnlineClaimsHelper claimsHelper = new MsOnlineClaimsHelper(URL, UserName, Password);

            try
            {
                _olist = _oweb.Lists.GetByTitle("ResourceAllocationDetails");
                camlquery = new CamlQuery();
                //camlquery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Variablename' /><Value Type='Text'>UAT StartDate</Value></Eq></Where></Query></View>";
                camlquery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>" + Id_traking + "</Value></Eq></Where></Query></View>";
                // ListItemCollection EmailvariablemappingItemsCollection = _olist.GetItems(camlquery);
                ListItemCollection listItems = _olist.GetItems(camlquery);
                ctx.Load(listItems, items => items.Include(item => item["Author"], item => item["ID"], item => item["Editor"], item => item["ResourceName"], item => item["EmployeeID"], item => item["ReasonForReject"]));
                ctx.ExecuteQuery();
                int Createdby_id = 0;
                string Createdby_names = "";
                int modifiedby_id = 0;
                string modifiedby_name = "";
                string EmpName = "";
                string Emp_ID = "";
                foreach (ListItem listItem in listItems)
                {
                    //Console.WriteLine("----------------------------------------");
                    //Console.WriteLine("Employee Name {0}", listItem["ResourceName"]);
                    EmpName = Convert.ToString(listItem["ResourceName"]);
                    Emp_ID = Convert.ToString(listItem["EmployeeID"]);
                    var Createdby_name = listItem["Author"] as FieldLookupValue;
                    if (Createdby_name != null)
                    {
                        Createdby_names = Createdby_name.LookupValue;
                        Createdby_id = Createdby_name.LookupId;
                    }

                    var modifiedby_name_ = listItem["Editor"] as FieldLookupValue;
                    if (Createdby_name != null)
                    {
                        modifiedby_name = modifiedby_name_.LookupValue;
                        modifiedby_id = modifiedby_name_.LookupId;
                    }
                    string RejCom = Convert.ToString(listItem["ReasonForReject"]);
                    bodyreturn = "<div> Hi " + Createdby_names + ",<br/><br/> Your request is being rejected by " + modifiedby_name + " for<br/><br/> " + EmpName + " " + Emp_ID + " because of below reason.<br/><br/><b> Reason of Rejection-</b> " + RejCom + "<br/><br/> Reach out to rmo@bcone.com for further clarifications.<br/><br/><font color='red'> This message is auto-generated and do not reply to this email.</font><br/><br/><b> Thanks & Regards,</b><br/><br/> RMO Team​​ </div><br/><br/>";

                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                SendMailsDatBaseProfile("pankaj.singh@e2eprojects.com", "uday.s@e2eprojects.com", ex.ToString(), "Error GetBody_EarlyRelaseReject");
            }
            return bodyreturn;
        }


        private static void ReplaceVariablevalue(ListItem itemsRRF, ClientContext ctx, Web _oweb, ref List _olist, ref CamlQuery camlquery, StringBuilder stringBuilder, List<string> results, string type, string FinalTo, string EventId)
        {
            JArray jarr = null;
            string ProjectManagerMail = "";
            MsOnlineClaimsHelper claimsHelper = new MsOnlineClaimsHelper(URL, UserName, Password);
            try
            {

                for (int i = 0; i < results.Count; i++)
                {
                    string variablename = results[i];
                    if (variablename == "CreatedBy")
                    {
                        var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/lists/getbytitle('RRF')/items?$select=NewAuthor/Title,NewAuthorId&$expand=NewAuthor&$filter=Id eq " + itemsRRF["ID"] + "");
                        request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                        request.Method = WebRequestMethods.Http.Get;
                        request.Accept = "application/json;odata=verbose";
                        // request.ContentType = "application/json;odata=verbose";
                        request.ContentLength = 0;

                        var securePassword = new SecureString();
                        foreach (char c in Password)
                        {
                            securePassword.AppendChar(c);
                        }
                        request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);
                        /* HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/web/lists/getbytitle('RRF')/items?$select=NewAuthor/Title,NewAuthorId&$expand=NewAuthor&$filter=Id eq " + itemsRRF["ID"] + "");
                         endpointRequest.Method = "GET";
                         //if (XML == false)
                         endpointRequest.Accept = "application/json;odata=verbose";
                         endpointRequest.UseDefaultCredentials = false;

                         endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                        HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                        WebResponse webResponse = request.GetResponse();
                        Stream webStream = webResponse.GetResponseStream();
                        StreamReader responseReader = new StreamReader(webStream);
                        string response = responseReader.ReadToEnd();
                        JObject jobj = JObject.Parse(response);
                        jarr = (JArray)jobj["d"]["results"];
                        JArray jarrPT = new JArray();
                        foreach (JObject j in jarr)
                        {
                            JObject jPT = new JObject();
                            string NewAuhtor = j["NewAuthor"]["Title"].ToString();
                            stringBuilder.Replace("&#123;" + variablename + "&#125;", NewAuhtor);
                        }

                    }
                    else
                    {
                        _olist = _oweb.Lists.GetByTitle("BconeEmailvariablemapping");
                        camlquery = new CamlQuery();
                        //camlquery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Variablename' /><Value Type='Text'>UAT StartDate</Value></Eq></Where></Query></View>";
                        camlquery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='VariableName' /><Value Type='Text'>" + variablename + "</Value></Eq></Where></Query></View>";
                        ListItemCollection EmailvariablemappingItemsCollection = _olist.GetItems(camlquery);
                        ctx.Load(EmailvariablemappingItemsCollection);
                        ctx.ExecuteQuery();


                         if (EmailvariablemappingItemsCollection.Count > 0)
                        {
                            int Flag = Convert.ToInt32(EmailvariablemappingItemsCollection[0]["Flag"]);
                            string FieldValue = Convert.ToString(EmailvariablemappingItemsCollection[0]["FieldValue"]);


                            string itemValue = "";
                            if (Flag == 1)
                            {
                                if (type == "Body")
                                {
                                    try
                                    {
                                        if (itemsRRF[FieldValue] != null && itemsRRF[FieldValue].ToString() != "")
                                        {
                                            if (FieldValue == "NewStartDate" || FieldValue == "NewEndDate" || FieldValue == "NewCreatedDate" || FieldValue == "NewReleaseDate" || FieldValue == "ReleaseDate" || FieldValue == "ExtensionDate" || FieldValue == "StartDate" || FieldValue == "EndDate" || FieldValue == "Modified" || FieldValue == "Author" || FieldValue == "Editor")
                                            {
                                                if (FieldValue == "NewCreatedDate")
                                                {
                                                    FieldValue = "RRFCreatedDate";
                                                }
                                                if ((itemsRRF[FieldValue].ToString().Contains(" PM") || itemsRRF[FieldValue].ToString().Contains("/") || (itemsRRF[FieldValue].ToString().Contains(" PM") || itemsRRF[FieldValue].ToString().Contains("-"))))
                                                {
                                                    itemValue = Convert.ToDateTime(itemsRRF[FieldValue]).ToShortDateString();
                                                }
                                                else if ((itemsRRF[FieldValue].ToString().Contains(" AM") || itemsRRF[FieldValue].ToString().Contains("/") || (itemsRRF[FieldValue].ToString().Contains(" AM") || itemsRRF[FieldValue].ToString().Contains("-"))))
                                                {
                                                    itemValue = Convert.ToDateTime(itemsRRF[FieldValue]).ToShortDateString();
                                                }
                                            }
                                            else
                                            {
                                                itemValue = itemsRRF[FieldValue].ToString();
                                            }
                                            stringBuilder.Replace("&#123;" + variablename + "&#125;", Convert.ToString(itemValue));
                                        }
                                    
                                    else
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", Convert.ToString(itemValue));
                                    }
                                    }
                                    catch (Exception ex)
                                    {

                                        ex.ToString();

                                    }

                                }
                                else
                                {
                                    if (itemsRRF[FieldValue] != null && itemsRRF[FieldValue].ToString() != "")
                                    {
                                        if (FieldValue == "NewStartDate" || FieldValue == "NewEndDate" || FieldValue == "NewCreatedDate" || FieldValue == "NewReleaseDate" || FieldValue == "ReleaseDate" || FieldValue == "ExtensionDate" || FieldValue == "StartDate" || FieldValue == "EndDate" || FieldValue == "Modified")
                                        {
                                            if (FieldValue == "NewCreatedDate")
                                            {
                                                FieldValue = "RRFCreatedDate";
                                            }
                                            if (itemsRRF[FieldValue].ToString().Contains(" PM") && itemsRRF[FieldValue].ToString().Contains("/") || (itemsRRF[FieldValue].ToString().Contains(" PM") && itemsRRF[FieldValue].ToString().Contains("-")))
                                            {
                                                itemValue = Convert.ToDateTime(itemsRRF[FieldValue]).ToShortDateString();
                                            }
                                            else if (itemsRRF[FieldValue].ToString().Contains(" AM") && itemsRRF[FieldValue].ToString().Contains("/") || (itemsRRF[FieldValue].ToString().Contains(" AM") && itemsRRF[FieldValue].ToString().Contains("-")))
                                            {
                                                itemValue = Convert.ToDateTime(itemsRRF[FieldValue]).ToShortDateString();
                                            }
                                        }

                                        else
                                        {
                                            itemValue = itemsRRF[FieldValue].ToString();
                                        }
                                        stringBuilder.Replace("{" + variablename + "}", Convert.ToString(itemValue));
                                    }
                                    else
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", Convert.ToString(itemValue));
                                    }
                                }
                            }
                            else if (Flag == 3)
                            {
                                var FilterValue = "";
                                
                                if (EventId == "54" && variablename == "PMContactNumber")
                                {
                                    FilterValue = "_api/ProjectData/Resources?$select=ResourceName,RoleBand,EmployeeRole,SubPractice,PrimarySkill,Skill,PhoneNumber&$filter=ResourceEmailAddress eq '" + ProjectManagerMail + "'";
                                }
                                else if (EventId == "54")
                                {
                                    FilterValue = "_api/ProjectData/Resources?$select=ResourceName,RoleBand,EmployeeRole,SubPractice,PrimarySkill,Skill,PhoneNumber&$filter=ResourceEmailAddress eq '" + ResourceEmailValue + "'";
                                }

                                else
                                {
                                    FilterValue = "_api/ProjectData/Resources?$select=ResourceName,RoleBand,EmployeeRole,SubPractice,PrimarySkill,Skill,PhoneNumber&$filter=EmployeeID eq '" + itemsRRF["EmployeeID"] + "'";
                                }
                                var request = (HttpWebRequest)WebRequest.Create(URL + FilterValue);
                                request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                                request.Method = WebRequestMethods.Http.Get;
                                request.Accept = "application/json;odata=verbose";
                                // request.ContentType = "application/json;odata=verbose";
                                request.ContentLength = 0;

                                var securePassword = new SecureString();
                                foreach (char c in Password)
                                {
                                    securePassword.AppendChar(c);
                                }
                                request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);
                                /* HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + FilterValue);
                                 endpointRequest.Method = "GET";
                                 //if (XML == false)
                                 endpointRequest.Accept = "application/json;odata=verbose";
                                 endpointRequest.UseDefaultCredentials = false;

                                 endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                                HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                                WebResponse webResponse = request.GetResponse();
                                Stream webStream = webResponse.GetResponseStream();
                                StreamReader responseReader = new StreamReader(webStream);
                                string response = responseReader.ReadToEnd();
                                JObject jobj = JObject.Parse(response);
                                jarr = (JArray)jobj["d"]["results"];
                                JArray jarrPT = new JArray();
                                foreach (JObject j in jarr)
                                {
                                    JObject jPT = new JObject();
                                    string RoleBand = j["RoleBand"].ToString();

                                    string EmployeeRole = j["EmployeeRole"].ToString();

                                    string SubPractice = j["SubPractice"].ToString();

                                    string PrimarySkill = j["PrimarySkill"].ToString();

                                    string Skill = j["Skill"].ToString();

                                    string PhoneNumber = j["PhoneNumber"].ToString();
                                  

                                    string AllocatedResourceName = j["ResourceName"].ToString();

                                    if (variablename == "ReleaseRoleBand")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", RoleBand);
                                    }
                                    else if (variablename == "ReleaseEmploymentRole")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", EmployeeRole);
                                    }
                                    else if (variablename == "ReleaseSubPractice")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", SubPractice);
                                    }
                                    else if (variablename == "ReleasePrimarySkill")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", PrimarySkill);
                                    }
                                    else if (variablename == "ReleaseSecondarySkill")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", Skill);
                                    }
                                    else if (variablename == "Billability")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", EmployeeRole);
                                    }
                                    else if (variablename == "PhoneNumber")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", PhoneNumber);
                                    }
                                    else if (variablename == "AllocatedResourceName")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", AllocatedResourceName);
                                    }
                                    else if (variablename == "PMContactNumber")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", PhoneNumber);
                                    }

                                    else
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", "");
                                    }



                                }
                            }
                            else if (Flag == 4)
                            {
                                var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/lists/getbytitle('RMOResourceAssignment')/items?$select=NewStartDate,NewEndDate,StartDate,EndDate,AllocationPercent,ProjectLoaction&$filter=RRFNumber eq '" + itemsRRF["RRFNO"] + "'");
                                request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                                request.Method = WebRequestMethods.Http.Get;
                                request.Accept = "application/json;odata=verbose";
                                // request.ContentType = "application/json;odata=verbose";
                                request.ContentLength = 0;

                                var securePassword = new SecureString();
                                foreach (char c in Password)
                                {
                                    securePassword.AppendChar(c);
                                }
                                request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);

                                /* HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/web/lists/getbytitle('RMOResourceAssignment')/items?$select=StartDate,EndDate,AllocationPercent,ProjectLoaction&$filter=RRFNumber eq '" + itemsRRF["RRFNO"] + "'");
                                 endpointRequest.Method = "GET";
                                 //if (XML == false)
                                 endpointRequest.Accept = "application/json;odata=verbose";
                                 endpointRequest.UseDefaultCredentials = false;

                                 endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                                HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                                WebResponse webResponse = request.GetResponse();
                                Stream webStream = webResponse.GetResponseStream();
                                StreamReader responseReader = new StreamReader(webStream);
                                string response = responseReader.ReadToEnd();
                                JObject jobj = JObject.Parse(response);
                                jarr = (JArray)jobj["d"]["results"];
                                JArray jarrPT = new JArray();
                                foreach (JObject j in jarr)
                                {
                                    JObject jPT = new JObject();
                                    string StartDate = j["StartDate"].ToString();
                                    StartDate= StartDate.Split(' ')[0];


                                    string EndDate = j["EndDate"].ToString();
                                    EndDate = EndDate.Split(' ')[0];

                                    string AllocationPer = j["AllocationPercent"].ToString();

                                    string ProjectLocation = j["ProjectLoaction"].ToString();

                                    if (variablename == "AllocationStartDate")
                                    {
                                        if (FieldValue == "NewStartDate" || FieldValue == "NewEndDate" || FieldValue == "NewCreatedDate" || FieldValue == "NewReleaseDate" || FieldValue == "ReleaseDate" || FieldValue == "ExtensionDate" || FieldValue == "StartDate" || FieldValue == "EndDate" || FieldValue == "Modified")
                                        {
                                            if ((j[FieldValue].ToString().Contains(" PM") && j[FieldValue].ToString().Contains("/") || (j[FieldValue].ToString().Contains(" PM") && j[FieldValue].ToString().Contains("-"))))
                                            {
                                                StartDate = Convert.ToDateTime(j[FieldValue]).ToShortDateString();
                                            }
                                            else if ((j[FieldValue].ToString().Contains(" AM") && j[FieldValue].ToString().Contains("/") || (j[FieldValue].ToString().Contains(" AM") && j[FieldValue].ToString().Contains("-"))))
                                            {
                                                StartDate = Convert.ToDateTime(j[FieldValue]).ToShortDateString();
                                            }
                                        }
                                        else
                                        {
                                            StartDate = "";
                                        }
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", StartDate);
                                    }
                                    else if (variablename == "AllocationEndDate")
                                    {
                                        if (FieldValue == "NewStartDate" || FieldValue == "NewEndDate" || FieldValue == "NewCreatedDate" || FieldValue == "NewReleaseDate" || FieldValue == "ReleaseDate" || FieldValue == "ExtensionDate" || FieldValue == "StartDate" || FieldValue == "EndDate" || FieldValue == "Modified")
                                        {
                                            if ((j[FieldValue].ToString().Contains(" PM") && j[FieldValue].ToString().Contains("/") || (j[FieldValue].ToString().Contains(" PM") && j[FieldValue].ToString().Contains("-"))))
                                            {
                                                EndDate = Convert.ToDateTime(j[FieldValue]).ToShortDateString();
                                            }
                                            else if ((j[FieldValue].ToString().Contains(" AM") && j[FieldValue].ToString().Contains("/") || (j[FieldValue].ToString().Contains(" AM") && j[FieldValue].ToString().Contains("-"))))
                                            {
                                                EndDate = Convert.ToDateTime(j[FieldValue]).ToShortDateString();
                                            }
                                        }
                                        else
                                        {
                                            EndDate = "";
                                        }
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", EndDate);
                                    }
                                    else if (variablename == "AllocationPercentage")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", AllocationPer);
                                    }
                                    else if (variablename == "ProjectLocation")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", ProjectLocation);
                                    }
                                    else
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", "");
                                    }

                                }
                            }
                            else if (Flag == 5 || Flag == 6)
                            {

                                if (Flag == 5)
                                {
                                    var request = (HttpWebRequest)WebRequest.Create(URL + "_api/ProjectData/Projects?$select=ProjectOwnerName,ClientPartner&$filter=ProjectCode eq '" + itemsRRF["AllocatedProjectCode"] + "'");
                                    request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                                    request.Method = WebRequestMethods.Http.Get;
                                    request.Accept = "application/json;odata=verbose";
                                    // request.ContentType = "application/json;odata=verbose";
                                    request.ContentLength = 0;

                                    var securePassword = new SecureString();
                                    foreach (char c in Password)
                                    {
                                        securePassword.AppendChar(c);
                                    }
                                    request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);
                                    /* HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/ProjectData/Projects?$select=ProjectOwnerName,ClientPartner&$filter=ProjectCode eq '" + itemsRRF["AllocatedProjectCode"] + "'");
                                     endpointRequest.Method = "GET";
                                     //if (XML == false)
                                     endpointRequest.Accept = "application/json;odata=verbose";
                                     endpointRequest.UseDefaultCredentials = false;

                                     endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                                    HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                                    WebResponse webResponse = request.GetResponse();
                                    Stream webStream = webResponse.GetResponseStream();
                                    StreamReader responseReader = new StreamReader(webStream);
                                    string response = responseReader.ReadToEnd();
                                    JObject jobj = JObject.Parse(response);
                                    jarr = (JArray)jobj["d"]["results"];
                                    JArray jarrPT = new JArray();
                                }
                                else if (Flag == 6)
                                {
                                    var request = (HttpWebRequest)WebRequest.Create(URL + "_api/ProjectData/Resources?$select=ReportingManager,BillableStatus&$filter=EmployeeID eq '" + itemsRRF["EmployeeID"] + "'");
                                    request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                                    request.Method = WebRequestMethods.Http.Get;
                                    request.Accept = "application/json;odata=verbose";
                                    // request.ContentType = "application/json;odata=verbose";
                                    request.ContentLength = 0;

                                    var securePassword = new SecureString();
                                    foreach (char c in Password)
                                    {
                                        securePassword.AppendChar(c);
                                    }
                                    request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);

                                    /* HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/ProjectData/Resources?$select=ReportingManager,BillableStatus&$filter=EmployeeID eq '" + itemsRRF["EmployeeID"] + "'");
                                     endpointRequest.Method = "GET";
                                     //if (XML == false)
                                     endpointRequest.Accept = "application/json;odata=verbose";
                                     endpointRequest.UseDefaultCredentials = false;

                                     endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                                    HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                                    WebResponse webResponse = request.GetResponse();
                                    Stream webStream = webResponse.GetResponseStream();
                                    StreamReader responseReader = new StreamReader(webStream);
                                    string response = responseReader.ReadToEnd();
                                    JObject jobj = JObject.Parse(response);
                                    jarr = (JArray)jobj["d"]["results"];
                                    JArray jarrPT = new JArray();
                                }



                                foreach (JObject j in jarr)
                                {
                                    JObject jPT = new JObject();

                                    if (Flag == 5)
                                    {
                                        if (variablename == "ExtensionProjectmanager")
                                        {
                                            string ProjectOwnerName = j["ProjectOwnerName"].ToString();
                                            stringBuilder.Replace("&#123;" + variablename + "&#125;", ProjectOwnerName);

                                        }
                                        else if (variablename == "ExtensionClientPartner")
                                        {
                                            string ClientPartner = j["ClientPartner"].ToString();
                                            if (ClientPartner != "")
                                            {
                                                ClientPartner = ClientPartner.Split('|')[1];
                                            }
                                            stringBuilder.Replace("&#123;" + variablename + "&#125;", ClientPartner);
                                        }
                                    }
                                    else if (Flag == 6)
                                    {
                                        if (variablename == "ExtensionBillableStatus")
                                        {
                                            string BillableStatus = j["BillableStatus"].ToString();
                                            stringBuilder.Replace("&#123;" + variablename + "&#125;", BillableStatus);
                                        }
                                        else if (variablename == "ExtensionReportingManager")
                                        {
                                            string ReportingManager = j["ReportingManager"].ToString();
                                            if (ReportingManager != "")
                                            {
                                                ReportingManager = ReportingManager.Split('|')[1];
                                            }
                                            stringBuilder.Replace("&#123;" + variablename + "&#125;", ReportingManager);
                                        }
                                    }

                                }
                            }
                            else
                            {
                                //var lookFieldvalue = (itemsRRF[FieldValue]) as FieldLookupValue;
                                if (type == "Body")
                                {
                                    if (itemsRRF[FieldValue] != null)
                                    {
                                        string Vlaue = ((Microsoft.SharePoint.Client.FieldLookupValue)itemsRRF[FieldValue]).LookupValue;
                                        if(EventId== "54" && variablename == "ProjectManager")
                                        {
                                            Vlaue = ((Microsoft.SharePoint.Client.FieldUserValue)((Microsoft.SharePoint.Client.FieldLookupValue)itemsRRF[FieldValue])).Email;
                                            ProjectManagerMail= Vlaue;
                                            
                                        }
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", Vlaue);
                                    }
                                    else
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", "");
                                    }

                                }
                                else
                                {
                                    if (itemsRRF[FieldValue] != null)
                                    {
                                        string Vlaue = ((Microsoft.SharePoint.Client.FieldLookupValue)itemsRRF[FieldValue]).LookupValue;
                                        stringBuilder.Replace("{" + variablename + "}", Vlaue);
                                    }
                                    else
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", "");
                                    }
                                }

                            }
                        }
                        else
                        {
                            stringBuilder.Replace("&#123;" + variablename + "&#125;", "");
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                ErrorFlag = "1";
                SendMailsDatBaseProfile("pankaj.singh@e2eprojects.com", "uday.s@e2eprojects.com", ex.ToString(), "Error ReplaceVariablevalue");
                Console.WriteLine(ex.Message);
            }
        }

        public static void sendEmailReleaseReject(ClientContext ctx, Web _oweb, ListItem itemsRRF, string EventId, string FinalTo, string FinalCc, string List_TrackingId)
        {
            try
            {
                //string authorName = Convert.ToString(((Microsoft.SharePoint.Client.FieldLookupValue)new System.Collections.Generic.Mscorlib_DictionaryDebugView<string, object>(itemsRRF.FieldValues).Items[57].Value).LookupValue);
                JArray jarr = null;
                MsOnlineClaimsHelper claimsHelper = new MsOnlineClaimsHelper(URL, UserName, Password);
                List _olist = _oweb.Lists.GetByTitle("BconeEmailConfiguration");
                CamlQuery camlquery = new CamlQuery();
                camlquery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='EventID' /><Value Type='Text'>" + EventId + "</Value></Eq></Where></Query></View>";
                ListItemCollection EmailConfigurationtItemsCollection = _olist.GetItems(camlquery);
                ctx.Load(EmailConfigurationtItemsCollection);
                ctx.ExecuteQuery();
                string FinalToEmailId = "";
                int m = 0;
                if (FinalTo != "")
                {
                    string[] newTo = FinalTo.Split(';');
                    foreach (string EmployeeID in newTo)
                    {
                        if (EmployeeID.ToString().Contains('@'))
                        {
                            if (m == 0)
                            {
                                FinalToEmailId = EmployeeID;
                            }
                            else
                            {
                                FinalToEmailId = FinalToEmailId + ";" + EmployeeID + ";";
                            }
                            m++;
                        }
                        else
                        {
                            var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                            request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                            request.Method = WebRequestMethods.Http.Get;
                            request.Accept = "application/json;odata=verbose";
                            // request.ContentType = "application/json;odata=verbose";
                            request.ContentLength = 0;

                            var securePassword = new SecureString();
                            foreach (char c in Password)
                            {
                                securePassword.AppendChar(c);
                            }
                            request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);
                            HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                            WebResponse webResponse = request.GetResponse();
                            Stream webStream = webResponse.GetResponseStream();
                            StreamReader responseReader = new StreamReader(webStream);
                            string response = responseReader.ReadToEnd();
                            JObject jobj = JObject.Parse(response);
                            jarr = (JArray)jobj["d"]["results"];
                            JArray jarrPT = new JArray();
                            foreach (JObject j in jarr)
                            {
                                JObject jPT = new JObject();
                                string emailId = j["Email"].ToString();
                                if (m == 0)
                                {
                                    FinalToEmailId = emailId;
                                }
                                else
                                {
                                    FinalToEmailId = FinalToEmailId + ";" + emailId + ";";
                                }
                                m++;
                            }
                        }
                    }
                }
                string FinalCcEmailId = "";
                if (FinalCc != "")
                {
                    string[] newCo = FinalCc.Split(';');
                    int o = 0;
                    foreach (string EmployeeID in newCo)
                    {
                        if (EmployeeID.ToString().Contains('@'))
                        {
                            if (o == 0)
                            {
                                FinalCcEmailId = EmployeeID;
                            }
                            else
                            {
                                FinalCcEmailId = FinalCcEmailId + ";" + EmployeeID + ";";
                            }
                            o++;
                        }
                        else
                        {
                            var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                            request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                            request.Method = WebRequestMethods.Http.Get;
                            request.Accept = "application/json;odata=verbose";
                            // request.ContentType = "application/json;odata=verbose";
                            request.ContentLength = 0;

                            var securePassword = new SecureString();
                            foreach (char c in Password)
                            {
                                securePassword.AppendChar(c);
                            }
                            request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);
                            HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                            WebResponse webResponse = request.GetResponse();
                            Stream webStream = webResponse.GetResponseStream();
                            StreamReader responseReader = new StreamReader(webStream);
                            string response = responseReader.ReadToEnd();
                            JObject jobj = JObject.Parse(response);
                            jarr = (JArray)jobj["d"]["results"];
                            JArray jarrPT = new JArray();
                            foreach (JObject j in jarr)
                            {
                                JObject jPT = new JObject();
                                string emailId = j["Email"].ToString();
                                if (o == 0)
                                {
                                    FinalCcEmailId = emailId;
                                }
                                else
                                {
                                    FinalCcEmailId = FinalCcEmailId + ";" + emailId + ";";
                                }
                                o++;
                            }
                        }
                    }
                }
                string body = Convert.ToString(EmailConfigurationtItemsCollection[0]["Body"]);
                string Subject = Convert.ToString(EmailConfigurationtItemsCollection[0]["Subject"]);
                string noHTMLBody = System.Text.RegularExpressions.Regex.Replace(body, @"<[^>]+>|&nbsp;", "").Trim();
                string noHTMLNormalisedBody = System.Text.RegularExpressions.Regex.Replace(noHTMLBody, @"\s{2,}", " ");
                string noHTMLSubject = System.Text.RegularExpressions.Regex.Replace(Subject, @"<[^>]+>|&nbsp;", "").Trim();
                string noHTMLNormalisednoHTMLSubject = System.Text.RegularExpressions.Regex.Replace(noHTMLSubject, @"\s{2,}", " ");
                StringBuilder stringBuilder = new StringBuilder(body);
                StringBuilder stringBuildersubject = new StringBuilder(Subject);

                List<string> Bodyvariable = ExtractFromString(noHTMLNormalisedBody, "&#123;", "&#125;");

                string final_body = GetBody_EarlyRelaseReject(List_TrackingId, ctx, _oweb, ref _olist, ref camlquery);

                // GetBody_EarlyRelaseReject(string Id_traking, ClientContext ctx, Web _oweb, ref List _olist, ref CamlQuery camlquery)

                List<string> Subjectvariable = ExtractFromString(noHTMLNormalisednoHTMLSubject, "{", "}");

                ReplaceVariablevalue(itemsRRF, ctx, _oweb, ref _olist, ref camlquery, stringBuildersubject, Subjectvariable, "Subject", FinalTo, EventId);

                /* List<string> Hyperlink = ExtractFromString(noHTMLNormalisedBody, "&lt;", "&gt;");

                 for (int i = 0; i < Hyperlink.Count; i++)
                 {
                     string variablename = Hyperlink[i];
                     stringBuilder.Replace("&lt;" + variablename + "&gt;", "<a style='color:#000000;face:Segoe UI Semibold, Calibri' href='" + RedirectURL + "'>Click Here</a>");

                 }*/



                string Body = string.Empty;

                Body = stringBuilder.ToString();
                Body = final_body;
                string finalsubject = stringBuildersubject.ToString();
                if (ErrorFlag != "1")
                {
                    //SqlCommand cmdExec = new SqlCommand("exec msdb.dbo.sp_send_dbmail @Profile_name=@Profile_name1," +
                    //                 "@recipients=@recipients1,@copy_recipients=@copy_recipients1,@subject=@subject1,@body=@body1,@body_format=@body_format1", con);
                    //if (con.State == ConnectionState.Closed)
                    //{
                    //    con.Open();
                    //}
                    //cmdExec.Parameters.AddWithValue("@Profile_name1", "RMO");
                    //cmdExec.Parameters.AddWithValue("@recipients1", FinalToEmailId);
                    //cmdExec.Parameters.AddWithValue("@subject1", finalsubject);
                    //cmdExec.Parameters.AddWithValue("@body1", Body);
                    //cmdExec.Parameters.AddWithValue("@copy_recipients1", FinalCcEmailId);
                    ////cmdExec.Parameters.AddWithValue("@blind_copy_recipients1", "Shweta.M@e2eprojects.com");

                    //cmdExec.Parameters.AddWithValue("@body_format1", "HTML");
                    //cmdExec.ExecuteNonQuery();
                    SendMailsDatBaseProfile(FinalToEmailId, FinalCcEmailId, Body, finalsubject);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                SendMailsDatBaseProfile("pankaj.singh@e2eprojects.com", "uday.s@e2eprojects.com", ex.ToString(), "Error sendEmailReleaseReject");
            }
        }


        public static void sendEmailRelease(ClientContext ctx, Web _oweb, ListItem itemsRRF, string EventId, string FinalTo, string FinalCc)
        {
            try
            {
                JArray jarr = null;
                MsOnlineClaimsHelper claimsHelper = new MsOnlineClaimsHelper(URL, UserName, Password);
                List _olist = _oweb.Lists.GetByTitle("BconeEmailConfiguration");
                CamlQuery camlquery = new CamlQuery();
                camlquery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='EventID' /><Value Type='Text'>" + EventId + "</Value></Eq></Where></Query></View>";
                ListItemCollection EmailConfigurationtItemsCollection = _olist.GetItems(camlquery);
                ctx.Load(EmailConfigurationtItemsCollection);
                ctx.ExecuteQuery();
                string FinalToEmailId = "";
                int m = 0;
                if (FinalTo != "")
                {
                    string[] newTo = FinalTo.Split(';');
                    foreach (string EmployeeID in newTo)
                    {
                        if (EmployeeID.ToString().Contains('@'))
                        {
                            if (m == 0)
                            {
                                FinalToEmailId = EmployeeID;
                            }
                            else
                            {
                                FinalToEmailId = FinalToEmailId + ";" + EmployeeID + ";";
                            }
                            m++;
                        }
                        else
                        {
                            var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                            request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                            request.Method = WebRequestMethods.Http.Get;
                            request.Accept = "application/json;odata=verbose";
                            // request.ContentType = "application/json;odata=verbose";
                            request.ContentLength = 0;

                            var securePassword = new SecureString();
                            foreach (char c in Password)
                            {
                                securePassword.AppendChar(c);
                            }
                            request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);

                            /* HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                             endpointRequest.Method = "GET";
                             //if (XML == false)
                             endpointRequest.Accept = "application/json;odata=verbose";
                             endpointRequest.UseDefaultCredentials = false;

                             endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                            HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                            WebResponse webResponse = request.GetResponse();
                            Stream webStream = webResponse.GetResponseStream();
                            StreamReader responseReader = new StreamReader(webStream);
                            string response = responseReader.ReadToEnd();
                            JObject jobj = JObject.Parse(response);
                            jarr = (JArray)jobj["d"]["results"];
                            JArray jarrPT = new JArray();
                            foreach (JObject j in jarr)
                            {
                                JObject jPT = new JObject();
                                string emailId = j["Email"].ToString();
                                if (m == 0)
                                {
                                    FinalToEmailId = emailId;
                                }
                                else
                                {
                                    FinalToEmailId = FinalToEmailId + ";" + emailId + ";";
                                }
                                m++;
                            }
                        }
                    }
                }
                string FinalCcEmailId = "";
                if (FinalCc != "")
                {
                    string[] newCo = FinalCc.Split(';');
                    int o = 0;
                    foreach (string EmployeeID in newCo)
                    {
                        if (EmployeeID.ToString().Contains('@'))
                        {
                            if (o == 0)
                            {
                                FinalCcEmailId = EmployeeID;
                            }
                            else
                            {
                                FinalCcEmailId = FinalCcEmailId + ";" + EmployeeID + ";";
                            }
                            o++;
                        }
                        else
                        {
                            var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                            request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                            request.Method = WebRequestMethods.Http.Get;
                            request.Accept = "application/json;odata=verbose";
                            // request.ContentType = "application/json;odata=verbose";
                            request.ContentLength = 0;

                            var securePassword = new SecureString();
                            foreach (char c in Password)
                            {
                                securePassword.AppendChar(c);
                            }
                            request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);

                            /* HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                             endpointRequest.Method = "GET";
                             //if (XML == false)
                             endpointRequest.Accept = "application/json;odata=verbose";
                             endpointRequest.UseDefaultCredentials = false;

                             endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                            HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                            WebResponse webResponse = request.GetResponse();
                            Stream webStream = webResponse.GetResponseStream();
                            StreamReader responseReader = new StreamReader(webStream);
                            string response = responseReader.ReadToEnd();
                            JObject jobj = JObject.Parse(response);
                            jarr = (JArray)jobj["d"]["results"];
                            JArray jarrPT = new JArray();
                            foreach (JObject j in jarr)
                            {
                                JObject jPT = new JObject();
                                string emailId = j["Email"].ToString();
                                if (o == 0)
                                {
                                    FinalCcEmailId = emailId;
                                }
                                else
                                {
                                    FinalCcEmailId = FinalCcEmailId + ";" + emailId + ";";
                                }
                                o++;
                            }
                        }
                    }
                }
                string body = Convert.ToString(EmailConfigurationtItemsCollection[0]["Body"]);
                string Subject = Convert.ToString(EmailConfigurationtItemsCollection[0]["Subject"]);
                string noHTMLBody = System.Text.RegularExpressions.Regex.Replace(body, @"<[^>]+>|&nbsp;", "").Trim();
                string noHTMLNormalisedBody = System.Text.RegularExpressions.Regex.Replace(noHTMLBody, @"\s{2,}", " ");
                var position = noHTMLNormalisedBody.Contains("myurl");

                string noHTMLSubject = System.Text.RegularExpressions.Regex.Replace(Subject, @"<[^>]+>|&nbsp;", "").Trim();
                string noHTMLNormalisednoHTMLSubject = System.Text.RegularExpressions.Regex.Replace(noHTMLSubject, @"\s{2,}", " ");
                StringBuilder stringBuilder = new StringBuilder(body);
                StringBuilder stringBuildersubject = new StringBuilder(Subject);

                List<string> Bodyvariable = ExtractFromString(noHTMLNormalisedBody, "&#123;", "&#125;");

                ReplaceVariablevalue(itemsRRF, ctx, _oweb, ref _olist, ref camlquery, stringBuilder, Bodyvariable, "Body", FinalTo, EventId);

                List<string> Subjectvariable = ExtractFromString(noHTMLNormalisednoHTMLSubject, "{", "}");

                ReplaceVariablevalue(itemsRRF, ctx, _oweb, ref _olist, ref camlquery, stringBuildersubject, Subjectvariable, "Subject", FinalTo, EventId);

                List<string> Hyperlink = ExtractFromString(noHTMLNormalisedBody, "&lt;", "&gt;");


                for (int i = 0; i < Hyperlink.Count; i++)
                {
                    string variablename = Hyperlink[i];
                    stringBuilder.Replace("&lt;" + variablename + "&gt;", "<a style='color:#000000;face:Segoe UI Semibold, Calibri' href='" + RedirectURL + "'>Click Here </a>");

                }

                string Body = string.Empty;

                Body = stringBuilder.ToString();

                string finalsubject = stringBuildersubject.ToString();
                if (ErrorFlag != "1")
                {
                    //SqlCommand cmdExec = new SqlCommand("exec msdb.dbo.sp_send_dbmail @Profile_name=@Profile_name1," +
                    //                 "@recipients=@recipients1,@copy_recipients=@copy_recipients1,@subject=@subject1,@body=@body1,@body_format=@body_format1", con);
                    //if (con.State == ConnectionState.Closed)
                    //{
                    //    con.Open();
                    //}
                    //cmdExec.Parameters.AddWithValue("@Profile_name1", "RMO");
                    //cmdExec.Parameters.AddWithValue("@recipients1", FinalToEmailId);
                    //cmdExec.Parameters.AddWithValue("@subject1", finalsubject);
                    //cmdExec.Parameters.AddWithValue("@body1", Body);
                    //cmdExec.Parameters.AddWithValue("@copy_recipients1", FinalCcEmailId); 
                    ////cmdExec.Parameters.AddWithValue("@blind_copy_recipients1", "Shweta.M@e2eprojects.com");

                    //cmdExec.Parameters.AddWithValue("@body_format1", "HTML");
                    //cmdExec.ExecuteNonQuery();
                    SendMailsDatBaseProfile(FinalToEmailId, FinalCcEmailId, Body, finalsubject);

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                SendMailsDatBaseProfile("pankaj.singh@e2eprojects.com", "uday.s@e2eprojects.com", ex.ToString(), "Error sendEmailRelease");
            }
        }


        public static void SendDynamicTableEmailRelease(ClientContext ctx, Web _oweb, string EventId, string FinalTo, string FinalCc, string TrackingId)
        {
            string semicolon = ";";
            string FinalToEmailId = "";
            string finalsubject = "";
            string FinalCcEmailId = "";
            string textBody = "";
            string newTrackingId = "";
            try
            {

                if (TrackingId.Contains(semicolon))
                {
                    newTrackingId = TrackingId.Split(';')[0];
                }
                else
                {
                    newTrackingId = TrackingId;
                }
                List _olistRRFMaster = _oweb.Lists.GetByTitle("ResourceAllocationDetails");
                var CamlQueryRRF = new CamlQuery() { ViewXml = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>" + newTrackingId + "</Value></Eq></Where></Query></View>" };
                ListItemCollection _olistItemsRRFCollection = _olistRRFMaster.GetItems(CamlQueryRRF);
                ctx.Load(_olistItemsRRFCollection);
                ctx.ExecuteQuery();
                foreach (ListItem itemsRRF in _olistItemsRRFCollection)
                {

                    JArray jarr = null;
                    MsOnlineClaimsHelper claimsHelper = new MsOnlineClaimsHelper(URL, UserName, Password);
                    List _olist = _oweb.Lists.GetByTitle("BconeEmailConfiguration");
                    CamlQuery camlqueryConfig = new CamlQuery();
                    camlqueryConfig.ViewXml = "<View><Query><Where><Eq><FieldRef Name='EventID' /><Value Type='Text'>" + EventId + "</Value></Eq></Where></Query></View>";
                    ListItemCollection EmailConfigurationtItemsCollection = _olist.GetItems(camlqueryConfig);
                    ctx.Load(EmailConfigurationtItemsCollection);
                    ctx.ExecuteQuery();

                    int k = 0;
                    if (FinalTo != "")
                    {
                        string[] newTo = FinalTo.Split(';');
                        foreach (string EmployeeID in newTo)
                        {
                            if (EmployeeID.ToString().Contains('@'))
                            {
                                if (k == 0)
                                {
                                    FinalToEmailId = EmployeeID;
                                }
                                else
                                {
                                    FinalToEmailId = FinalToEmailId + ";" + EmployeeID + ";";
                                }
                                k++;
                            }
                            else
                            {

                                var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                                request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                                request.Method = WebRequestMethods.Http.Get;
                                request.Accept = "application/json;odata=verbose";
                                request.ContentType = "application/json;odata=verbose";
                                request.ContentLength = 0;

                                var securePassword = new SecureString();
                                foreach (char c in Password)
                                {
                                    securePassword.AppendChar(c);
                                }
                                request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);

                                /*   using (var response = (HttpWebResponse)request.GetResponse())
                                   {
                                       using (var streamReader = new StreamReader(response.GetResponseStream()))
                                       {
                                           var content = streamReader.ReadToEnd();
                                           var t = JToken.Parse(content);
                                           return t["d"]["GetContextWebInformation"]["FormDigestValue"].ToString();
                                       }
                                   }


                                   HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                                   endpointRequest.Method = "GET";
                                   //if (XML == false)
                                   endpointRequest.Accept = "application/json;odata=verbose";
                                   endpointRequest.UseDefaultCredentials = false;

                                   endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                                HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                                WebResponse webResponse = request.GetResponse();
                                Stream webStream = webResponse.GetResponseStream();
                                StreamReader responseReader = new StreamReader(webStream);
                                string response = responseReader.ReadToEnd();
                                JObject jobj = JObject.Parse(response);
                                jarr = (JArray)jobj["d"]["results"];
                                JArray jarrPT = new JArray();
                                foreach (JObject j in jarr)
                                {
                                    JObject jPT = new JObject();
                                    string emailId = j["Email"].ToString();
                                    if (k == 0)
                                    {
                                        FinalToEmailId = emailId;
                                    }
                                    else
                                    {
                                        FinalToEmailId = FinalToEmailId + ";" + emailId + ";";
                                    }
                                    k++;
                                }
                            }
                        }
                    }

                    if (FinalCc != "")
                    {
                        string[] newCo = FinalCc.Split(';');
                        int l = 0;
                        foreach (string EmployeeID in newCo)
                        {
                            if (EmployeeID.ToString().Contains('@'))
                            {
                                if (l == 0)
                                {
                                    FinalCcEmailId = EmployeeID;
                                }
                                else
                                {
                                    FinalCcEmailId = FinalCcEmailId + ";" + EmployeeID + ";";
                                }
                                l++;
                            }
                            else
                            {
                                var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                                request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                                request.Method = WebRequestMethods.Http.Get;
                                request.Accept = "application/json;odata=verbose";
                                // request.ContentType = "application/json;odata=verbose";
                                request.ContentLength = 0;

                                var securePassword = new SecureString();
                                foreach (char c in Password)
                                {
                                    securePassword.AppendChar(c);
                                }
                                request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);


                                /*  HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/web/siteusers?$select=Email&$filter=Id eq " + EmployeeID + "");
                                  endpointRequest.Method = "GET";
                                  //if (XML == false)
                                  endpointRequest.Accept = "application/json;odata=verbose";
                                  endpointRequest.UseDefaultCredentials = false;

                                  endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                                HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                                WebResponse webResponse = request.GetResponse();
                                Stream webStream = webResponse.GetResponseStream();
                                StreamReader responseReader = new StreamReader(webStream);
                                string response = responseReader.ReadToEnd();
                                JObject jobj = JObject.Parse(response);
                                jarr = (JArray)jobj["d"]["results"];
                                JArray jarrPT = new JArray();
                                foreach (JObject j in jarr)
                                {
                                    JObject jPT = new JObject();
                                    string emailId = j["Email"].ToString();
                                    if (l == 0)
                                    {
                                        FinalCcEmailId = emailId;
                                    }
                                    else
                                    {
                                        FinalCcEmailId = FinalCcEmailId + ";" + emailId + ";";
                                    }
                                    l++;
                                }
                            }
                        }
                    }

                    string body = Convert.ToString(EmailConfigurationtItemsCollection[0]["Body"]);
                    string Subject = Convert.ToString(EmailConfigurationtItemsCollection[0]["Subject"]);
                    string noHTMLBody = System.Text.RegularExpressions.Regex.Replace(body, @"<[^>]+>|&nbsp;", "").Trim();
                    string noHTMLNormalisedBody = System.Text.RegularExpressions.Regex.Replace(noHTMLBody, @"\s{2,}", " ");
                    string noHTMLSubject = System.Text.RegularExpressions.Regex.Replace(Subject, @"<[^>]+>|&nbsp;", "").Trim();
                    string noHTMLNormalisednoHTMLSubject = System.Text.RegularExpressions.Regex.Replace(noHTMLSubject, @"\s{2,}", " ");
                    StringBuilder stringBuilder = new StringBuilder(body);
                    StringBuilder stringBuildersubject = new StringBuilder(Subject);

                    List<string> Bodyvariable = ExtractFromString(noHTMLNormalisedBody, "&#123;", "&#125;");

                    ReplaceVariablevalue1(itemsRRF, ctx, _oweb, ref _olist, ref camlqueryConfig, stringBuilder, Bodyvariable, "Body", FinalTo, EventId);

                    List<string> Subjectvariable = ExtractFromString(noHTMLNormalisednoHTMLSubject, "{", "}");

                    ReplaceVariablevalue(itemsRRF, ctx, _oweb, ref _olist, ref camlqueryConfig, stringBuildersubject, Subjectvariable, "Subject", FinalTo, EventId);

                    List<string> Hyperlink = ExtractFromString(noHTMLNormalisedBody, "&lt;", "&gt;");

                    for (int i = 0; i < Hyperlink.Count; i++)
                    {
                        string variablename = Hyperlink[i];
                        stringBuilder.Replace("&lt;" + variablename + "&gt;", "<a style='color:#000000;face:Segoe UI Semibold, Calibri' href='" + ReleaseRedirectURL + "'>Click Here </a>");

                    }

                    string Body = string.Empty;

                    Body = stringBuilder.ToString();

                    finalsubject = stringBuildersubject.ToString();



                    textBody = "<br />" + Body + "<br />" +
                    "<table class='MsoTableGrid' cellspacing='0' cellpadding='0' width='100%' border='1' style='table-layout: fixed;border-width: medium; border-style: none; border-color: initial; width: 559.7pt; margin: auto auto auto -0.25pt;'>" +
    "<tr>" +
    "<td valign='top' width='37' style='border-width: 1pt; border-style: solid; border-color: windowtext; width: 27.8pt; background: #2f5496; padding: 0in 5.4pt;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Sr No</font></font></span></p>" +
    "</td>" +
    "<td valign='top' width='72' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 54.35pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Employee Code</font></font></span></p>" +
    "</td>" +
    "<td valign='top' width='88' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 66.35pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Employee Name</font></font></span></p>" +
    "</td>" +
    "<td valign='top' width='150px' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 100px; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Project Code</font></font></span></p>" +
    "</td>" +
    "<td valign='top' width='200px' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 180px; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;width:150px'><font face='Segoe UI'><font size='2'>Project Name</font></font></span></p>" +
    "</td>" +
    "<td valign='top' width='77' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 57.65pt; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Allocation %</font></font></span></p>" +
    "</td>" +
    "<td valign='top' width='62' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 80px; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Original Start Date</font></font></span></p>" +
    "</td>" +
    "<td valign='top' width='73' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 80px; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Original End Date</font></font></span></p>" +
    "</td>" +
    "<td valign='top' width='73' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 75px; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Revised Allocation %</font></font></span></p>" +
    "</td>"
    +
    "<td valign='top' width='73' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 100px; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Revised Start Date</font></font></span></p>" +
    "</td>"
    +
    "<td valign='top' width='73' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 100px; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
    "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
    "<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Revised End Date</font></font></span></p>" +
    "</td>";

                     if (itemsRRF["ResourceType"].ToString() != "Resource Extension")
                    {
                        textBody += "<td valign='top' width='73' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 100px; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
                        "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
                        "<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'>Early Release Reason </font></font></span></p>" +
                        "</td>"

                        +
                        "<td valign='top' width='73' style='border-top: 1pt solid windowtext; border-right: 1pt solid windowtext; width: 100px; background: #2f5496; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0;'>" +
                        "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
                        "<span style='color: #ffffff;'><font face='Segoe UI'><font size='2'></font>Feedback</font></span></p>" +
                        "</td>";
                    }
                    else
                    {
                        // textBody += "</tr>";
                    }




                    if (TrackingId.Contains(semicolon))
                    {
                        int d = 1;
                        string[] ArrayTrackingId = TrackingId.Split(';');
                        foreach (var FinalTrackind in ArrayTrackingId)
                        {

                            string EmployeeCodeValue = "";
                            string EmaployeeName = "";
                            string ProjectCode = "";
                            string ProjectName = "";
                            string Alocation = "";
                            string StartDate = "";
                            string EndDate = "";
                            string RevEndDate = "";
                            string Early_Release_Reason = "";
                            string Feedback = "";
                            List _olistResourceAll = _oweb.Lists.GetByTitle("ResourceAllocationDetails");
                            CamlQuery camlquery = new CamlQuery();
                            camlquery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>" + FinalTrackind + "</Value></Eq></Where></Query></View>";
                            ListItemCollection ReleaseAllConfigurationtItemsCollection = _olistResourceAll.GetItems(camlquery);
                            ctx.Load(ReleaseAllConfigurationtItemsCollection);
                            ctx.ExecuteQuery();
                            if (ReleaseAllConfigurationtItemsCollection.Count > 0)
                            {

                                if (itemsRRF["EmployeeID"] != null)
                                {
                                    EmployeeCodeValue = Convert.ToString(ReleaseAllConfigurationtItemsCollection[0]["EmployeeID"]);
                                    if (EmployeeCodeValue.IndexOf('\t') > -1)
                                    {
                                        EmployeeCodeValue = EmployeeCodeValue.Replace("\t", "");
                                    }
                                }
                                if (ReleaseAllConfigurationtItemsCollection[0]["ResourceName"] != null && ReleaseAllConfigurationtItemsCollection[0]["ResourceName"].ToString() != "")
                                    EmaployeeName = Convert.ToString(ReleaseAllConfigurationtItemsCollection[0]["ResourceName"]);
                                if (ReleaseAllConfigurationtItemsCollection[0]["AllocatedProjectCode"] != null)
                                    ProjectCode = Convert.ToString(ReleaseAllConfigurationtItemsCollection[0]["AllocatedProjectCode"]);
                                if (ReleaseAllConfigurationtItemsCollection[0]["ProjectName"] != null)
                                    ProjectName = Convert.ToString(ReleaseAllConfigurationtItemsCollection[0]["ProjectName"]);
                                if (ReleaseAllConfigurationtItemsCollection[0]["AllocationPercentage"] != null)
                                    Alocation = Convert.ToString(ReleaseAllConfigurationtItemsCollection[0]["AllocationPercentage"]);
                                if (ReleaseAllConfigurationtItemsCollection[0]["NewStartDate"] != null)
                                    StartDate = Convert.ToDateTime(ReleaseAllConfigurationtItemsCollection[0]["NewStartDate"]).ToShortDateString();
                                if (ReleaseAllConfigurationtItemsCollection[0]["NewEndDate"] != null && ReleaseAllConfigurationtItemsCollection[0]["NewEndDate"].ToString() != "null" && ReleaseAllConfigurationtItemsCollection[0]["NewEndDate"].ToString() != "")
                                    EndDate = Convert.ToDateTime(ReleaseAllConfigurationtItemsCollection[0]["NewEndDate"]).ToShortDateString();

                                if (ReleaseAllConfigurationtItemsCollection[0]["ReasonForRelease"] != null)
                                {
                                    Early_Release_Reason = Convert.ToString(ReleaseAllConfigurationtItemsCollection[0]["ReasonForRelease"]);
                                }
                                else
                                {
                                    Early_Release_Reason = "None";
                                }
                                if (ReleaseAllConfigurationtItemsCollection[0]["ProjectEndFeedBack"] != null)
                                {
                                    Feedback = Convert.ToString(ReleaseAllConfigurationtItemsCollection[0]["ProjectEndFeedBack"]);
                                }
                                else
                                {
                                    Feedback = "None";
                                }

                                if (EventId == "64")
                                {
                                    if (ReleaseAllConfigurationtItemsCollection[0]["NewReleaseDate"] != null)
                                        RevEndDate = Convert.ToDateTime(ReleaseAllConfigurationtItemsCollection[0]["NewReleaseDate"]).ToShortDateString();

                                }
                                else if (EventId == "66")
                                {
                                    if (ReleaseAllConfigurationtItemsCollection[0]["NewExtensionDate"] != null)
                                        RevEndDate = Convert.ToDateTime(ReleaseAllConfigurationtItemsCollection[0]["NewExtensionDate"]).ToShortDateString();
                                }
                                textBody += "<tr><td valign='top' width='37' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 27.8pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: 1pt solid windowtext; background-color: transparent;'>" +
                "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
                "<span style='color: #1a1a1a;'><span>" + d + "</span><br/></span></p>" +
                "</td>" +
                "<td valign='top' width='72' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 54.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
                "<span style='color: #1a1a1a;'><span>" + EmployeeCodeValue + "​</span><br/></span></td>" +
                "<td valign='top' width='88' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 66.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
                "<span style='color: #1a1a1a;'><span>" + EmaployeeName + "</span><br/></span></td>" +
                "<td valign='top' width='25' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 18.85pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
                "<span style='color: #1a1a1a;'><span> " + ProjectCode + "</span><br/></span></td>" +
                "<td valign='top' width='61' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 45.75pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
                "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
                "<span style='color: #1a1a1a;'><span> " + ProjectName + "</span><br/></span></p>" +
                "</td>" +
                "<td valign='top' width='77' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 57.65pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
                "<span style='color: #1a1a1a;'>" + Alocation + "</span>" +
                "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
                "</p>" +
                "</td>" +
                "<td valign='top' width='62' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 46.6pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
                "<span style='color: #1a1a1a;'><span> " + StartDate + "</span><br/></span></td>" +
                "<td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
                "<span style='color: #1a1a1a;'><span> " + EndDate + "</span><br/></span></td>" +
                "<td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
                "<span style='color: #1a1a1a;'><span> " + Alocation + "</span><br/></span></td>" +
                "<td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
                "<span style='color: #1a1a1a;'><span> " + StartDate + "</span><br/></span></td>" +
                "<td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
                "<span style='color: #1a1a1a;'><span> " + RevEndDate + "</span><br/></span></td>";

                                if (ReleaseAllConfigurationtItemsCollection[0]["ResourceType"].ToString() != "Resource Extension")
                                {

                                    textBody += "<td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
                                    "<span style='color: #1a1a1a;'><span> " + Early_Release_Reason + "</span><br/></span></td>" +
                                    "<td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
                                    "<span style='color: #1a1a1a;'><span> " + Feedback + "</span><br/></span></td></tr>";
                                }
                                else
                                {
                                    textBody += "</tr>";
                                }

                            }
                            d++;
                        }

                        textBody += "</table><br />" +
                        "</tr><br/><font size='2'>" +
                        "<font color='#cc6600'>This message is auto-generated and do not reply to this email.​<br/><br/></font><strong><font color='#000000'>Thanks &amp; Regards,</font></strong><font color='#000000'></font><br/>RMO Team</font></p><br /><br />";


                    }
                    else
                    {
                        int a = 1;
                        string EmployeeCodeValue = "";
                        string EmaployeeName = "";
                        string ProjectCode = "";
                        string ProjectName = "";
                        string Alocation = "";
                        string StartDate = "";
                        string EndDate = "";
                        string RevEndDate = "";

                        string Feedback = "";
                        string Early_Release_Reason = "";

                        if (itemsRRF["EmployeeID"] != null)
                        {
                            EmployeeCodeValue = Convert.ToString(itemsRRF["EmployeeID"]);
                            if (EmployeeCodeValue.IndexOf('\t') > -1)
                            {
                                EmployeeCodeValue = EmployeeCodeValue.Replace("\t", "");
                            }
                        }
                        if (itemsRRF["ResourceName"] != null && itemsRRF["ResourceName"].ToString() != "")
                            EmaployeeName = Convert.ToString(itemsRRF["ResourceName"]);
                        if (itemsRRF["AllocatedProjectCode"] != null)
                            ProjectCode = Convert.ToString(itemsRRF["AllocatedProjectCode"]);
                        if (itemsRRF["ProjectName"] != null)
                            ProjectName = Convert.ToString(itemsRRF["ProjectName"]);
                        if (itemsRRF["AllocationPercentage"] != null)
                            Alocation = Convert.ToString(itemsRRF["AllocationPercentage"]);
                        if (itemsRRF["NewStartDate"] != null)
                            StartDate = Convert.ToDateTime(itemsRRF["NewStartDate"]).ToShortDateString();
                        if (itemsRRF["NewEndDate"] != null && itemsRRF["NewEndDate"].ToString() != "null" && itemsRRF["NewEndDate"].ToString() != "")
                            EndDate = Convert.ToDateTime(itemsRRF["NewEndDate"]).ToShortDateString();

                        if (itemsRRF["ReasonForRelease"] != null)
                        {
                            Early_Release_Reason = Convert.ToString(itemsRRF["ReasonForRelease"]);
                        }
                        else
                        {
                            Early_Release_Reason = "None";
                        }
                        if (itemsRRF["ProjectEndFeedBack"] != null)
                        {
                            Feedback = Convert.ToString(itemsRRF["ProjectEndFeedBack"]);
                        }
                        else
                        {
                            Feedback = "None";
                        }

                        if (EventId == "64")
                        {
                            if (itemsRRF["NewReleaseDate"] != null && itemsRRF["NewReleaseDate"].ToString() != "")
                                RevEndDate = Convert.ToDateTime(itemsRRF["NewReleaseDate"]).ToShortDateString();
                        }
                        else if (EventId == "66")
                        {
                            if (itemsRRF["NewExtensionDate"] != null && itemsRRF["NewExtensionDate"].ToString() != "")
                                RevEndDate = Convert.ToDateTime(itemsRRF["NewExtensionDate"]).ToShortDateString();
                        }

                        textBody += "<tr><td valign='top' width='37' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 27.8pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: 1pt solid windowtext; background-color: transparent;'>" +
        "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
        "<span style='color: #1a1a1a;'><span>" + a + "</span><br/></span></p>" +
        "</td>" +
        "<td valign='top' width='72' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 54.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
        "<span style='color: #1a1a1a;'><span>" + EmployeeCodeValue + "​</span><br/></span></td>" +
        "<td valign='top' width='88' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 66.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
        "<span style='color: #1a1a1a;'><span>" + EmaployeeName + "</span><br/></span></td>" +
        "<td valign='top' width='25' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 18.85pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
        "<span style='color: #1a1a1a;'><span> " + ProjectCode + "</span><br/></span></td>" +
        "<td valign='top' width='61' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 45.75pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
        "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
        "<span style='color: #1a1a1a;'><span> " + ProjectName + "</span><br/></span></p>" +
        "</td>" +
        "<td valign='top' width='77' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 57.65pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
        "<span style='color: #1a1a1a;'>" + Alocation + "</span>" +
        "<p class='MsoNormal' style='margin: 0in 0in 0pt; line-height: normal;'>" +
        "</p>" +
        "</td>" +
        "<td valign='top' width='62' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 46.6pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
        "<span style='color: #1a1a1a;'><span> " + StartDate + "</span><br/></span></td>" +
        "<td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
        "<span style='color: #1a1a1a;'><span> " + EndDate + "</span><br/></span></td>" +
        "<td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
        "<span style='color: #1a1a1a;'><span> " + Alocation + "</span><br/></span></td>" +
        "<td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
        "<span style='color: #1a1a1a;'><span> " + StartDate + "</span><br/></span></td>" +
        "<td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'>" +
        "<span style='color: #1a1a1a;'><span> " + RevEndDate + "</span><br/></span></td>";


                        if (itemsRRF["ResourceType"].ToString() != "Resource Extension")
                        {
                            textBody += "<td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'><span style='color: #1a1a1a;'><span> " + Early_Release_Reason + "</span><br/></span></td><td valign='top' width='73' style='border-top: #f0f0f0; border-right: 1pt solid windowtext; width: 32.35pt; border-bottom: 1pt solid windowtext; padding: 0in 5.4pt; border-left: #f0f0f0; background-color: transparent;'><span style='color: #1a1a1a;'><span> " + Feedback + "</span><br/></span></td></tr>";
                        }
                        else
                        {
                            textBody += "</tr>";
                        }


                        a++;



                        textBody += "</table><br />" +
        "</tr><br/><font size='2'>" +
                 "<font color='#cc6600'>This message is auto-generated and do not reply to this email.​<br/><br/></font><strong><font color='#000000'>Thanks &amp; Regards,</font></strong><font color='#000000'></font><br/>RMO Team</font></p><br /><br />";


                    }
                }
            }
            catch (Exception ex)
            {
                ErrorFlag = "1";
                // throw;
                SendMailsDatBaseProfile("pankaj.singh@e2eprojects.com", "uday.s@e2eprojects.com", ex.ToString(), "Error SendDynamicTableEmailRelease");
            }

            if (ErrorFlag != "1")
            {
                //try
                //{
                //    SqlCommand cmdExec = new SqlCommand("exec msdb.dbo.sp_send_dbmail @Profile_name=@Profile_name1," +
                //                                    "@recipients=@recipients1,@copy_recipients=@copy_recipients1,@subject=@subject1,@body=@body1,@body_format=@body_format1", con);
                //    if (con.State == ConnectionState.Closed)
                //    {
                //        con.Open();
                //    }
                //    //FinalToEmailId = "uday.s@e2eprojects.com";
                //    //FinalCcEmailId = "pankaj.singh@e2eprojects.com";
                //    cmdExec.Parameters.AddWithValue("@Profile_name1", "RMO");
                //    cmdExec.Parameters.AddWithValue("@recipients1", FinalToEmailId);
                //    cmdExec.Parameters.AddWithValue("@subject1", finalsubject);
                //    cmdExec.Parameters.AddWithValue("@body1", textBody);
                //    cmdExec.Parameters.AddWithValue("@copy_recipients1", FinalCcEmailId);
                //    //cmdExec.Parameters.AddWithValue("@blind_copy_recipients1s", "Getha.sk@e2eprojects.com");
                //    cmdExec.Parameters.AddWithValue("@body_format1", "HTML");
                //    cmdExec.ExecuteNonQuery();
                //}
                //catch (Exception ex)
                //{
                //    Console.WriteLine(ex.Message);
                //    SendMailsDatBaseProfile("pankaj.singh@e2eprojects.com", "uday.s@e2eprojects.com", ex.ToString(), "Error SendDynamicTableEmailRelease");
                //    // throw;
                //}
                SendMailsDatBaseProfile(FinalToEmailId, FinalCcEmailId, textBody, finalsubject);

            }


        }


        private static void ReplaceVariablevalue1(ListItem itemsRRF, ClientContext ctx, Web _oweb, ref List _olist, ref CamlQuery camlquery, StringBuilder stringBuilder, List<string> results, string type, string FinalTo,string _EventId)
        {
            JArray jarr = null;
            MsOnlineClaimsHelper claimsHelper = new MsOnlineClaimsHelper(URL, UserName, Password);
            try
            {

                for (int i = 0; i < results.Count; i++)
                {
                    string variablename = results[i];
                    if (variablename == "CreatedBy")
                    {

                        var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/lists/getbytitle('ResourceAllocationDetails')/items?$select=NewAuthor/Title,NewAuthorId&$expand=NewAuthor&$filter=Id eq " + itemsRRF["ID"] + "");
                        request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                        request.Method = WebRequestMethods.Http.Get;
                        request.Accept = "application/json;odata=verbose";
                        request.ContentLength = 0;
                        var securePassword = new SecureString();
                        foreach (char c in Password)
                        {
                            securePassword.AppendChar(c);
                        }
                        request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);

                        HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                        WebResponse webResponse = request.GetResponse();
                        Stream webStream = webResponse.GetResponseStream();
                        StreamReader responseReader = new StreamReader(webStream);
                        string response = responseReader.ReadToEnd();
                        JObject jobj = JObject.Parse(response);
                        jarr = (JArray)jobj["d"]["results"];
                        JArray jarrPT = new JArray();
                        foreach (JObject j in jarr)
                        {
                            JObject jPT = new JObject();
                            string NewAuhtor = j["NewAuthor"]["Title"].ToString();
                            stringBuilder.Replace("&#123;" + variablename + "&#125;", NewAuhtor);
                        }

                    }
                    else
                    {
                        _olist = _oweb.Lists.GetByTitle("BconeEmailvariablemapping");
                        camlquery = new CamlQuery();
                        //camlquery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Variablename' /><Value Type='Text'>UAT StartDate</Value></Eq></Where></Query></View>";
                        camlquery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='VariableName' /><Value Type='Text'>" + variablename + "</Value></Eq></Where></Query></View>";
                        ListItemCollection EmailvariablemappingItemsCollection = _olist.GetItems(camlquery);
                        ctx.Load(EmailvariablemappingItemsCollection);
                        ctx.ExecuteQuery();


                        if (EmailvariablemappingItemsCollection.Count > 0)
                        {
                            int Flag = Convert.ToInt32(EmailvariablemappingItemsCollection[0]["Flag"]);
                            string FieldValue = Convert.ToString(EmailvariablemappingItemsCollection[0]["FieldValue"]);


                            string itemValue = "";
                            if (Flag == 1)
                            {
                                if (type == "Body")
                                {
                                    if(FieldValue == "MailTo")
                                    {

                                    }
                                    if (itemsRRF[FieldValue] != null && itemsRRF[FieldValue].ToString() != "")
                                    {
                                        if (FieldValue == "NewStartDate" || FieldValue == "NewEndDate" || FieldValue == "NewCreatedDate" || FieldValue == "NewReleaseDate" || FieldValue == "ReleaseDate" || FieldValue == "ExtensionDate" || FieldValue == "StartDate" || FieldValue == "EndDate" || FieldValue == "Modified")
                                        {
                                            if (FieldValue == "NewCreatedDate")
                                            {
                                                FieldValue = "RRFCreatedDate";
                                            }
                                            if ((itemsRRF[FieldValue].ToString().Contains(" PM") && itemsRRF[FieldValue].ToString().Contains("/") || (itemsRRF[FieldValue].ToString().Contains(" PM") && itemsRRF[FieldValue].ToString().Contains("-"))))
                                            {
                                                itemValue = Convert.ToDateTime(itemsRRF[FieldValue]).ToShortDateString();
                                            }
                                            else if ((itemsRRF[FieldValue].ToString().Contains(" AM") && itemsRRF[FieldValue].ToString().Contains("/") || (itemsRRF[FieldValue].ToString().Contains(" AM") && itemsRRF[FieldValue].ToString().Contains("-"))))
                                            {
                                                itemValue = Convert.ToDateTime(itemsRRF[FieldValue]).ToShortDateString();
                                            }
                                        }
                                        else
                                        {
                                            itemValue = itemsRRF[FieldValue].ToString();
                                        }
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", Convert.ToString(itemValue));
                                    }
                                    else
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", Convert.ToString(itemValue));
                                    }

                                }
                                else
                                {
                                    if (itemsRRF[FieldValue] != null && itemsRRF[FieldValue].ToString() != "")
                                    {
                                        if (FieldValue == "NewStartDate" || FieldValue == "NewEndDate" || FieldValue == "NewCreatedDate" || FieldValue == "NewReleaseDate" || FieldValue == "ReleaseDate" || FieldValue == "ExtensionDate" || FieldValue == "StartDate" || FieldValue == "EndDate" || FieldValue == "Modified")
                                        {
                                            if (FieldValue == "NewCreatedDate")
                                            {
                                                FieldValue = "RRFCreatedDate";
                                            }
                                            if (itemsRRF[FieldValue].ToString().Contains(" PM") && itemsRRF[FieldValue].ToString().Contains("/") || (itemsRRF[FieldValue].ToString().Contains(" PM") && itemsRRF[FieldValue].ToString().Contains("-")))
                                            {
                                                itemValue = Convert.ToDateTime(itemsRRF[FieldValue]).ToShortDateString();
                                            }
                                            else if (itemsRRF[FieldValue].ToString().Contains(" AM") && itemsRRF[FieldValue].ToString().Contains("/") || (itemsRRF[FieldValue].ToString().Contains(" AM") && itemsRRF[FieldValue].ToString().Contains("-")))
                                            {
                                                itemValue = Convert.ToDateTime(itemsRRF[FieldValue]).ToShortDateString();
                                            }
                                        }
                                        else
                                        {
                                            itemValue = itemsRRF[FieldValue].ToString();
                                        }
                                        stringBuilder.Replace("{" + variablename + "}", Convert.ToString(itemValue));
                                    }
                                    else
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", Convert.ToString(itemValue));
                                    }
                                }
                            }
                            else if (Flag == 3)
                            {
                                var request = (HttpWebRequest)WebRequest.Create(URL + "_api/ProjectData/Resources?$select=RoleBand,ResourceName,EmployeeRole,SubPractice,PrimarySkill,Skill&$filter=EmployeeID eq '" + itemsRRF["EmployeeID"] + "'");
                                request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                                request.Method = WebRequestMethods.Http.Get;
                                request.Accept = "application/json;odata=verbose";
                                // request.ContentType = "application/json;odata=verbose";
                                request.ContentLength = 0;

                                var securePassword = new SecureString();
                                foreach (char c in Password)
                                {
                                    securePassword.AppendChar(c);
                                }
                                request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);

                                /*  HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/ProjectData/Resources?$select=RoleBand,EmployeeRole,SubPractice,PrimarySkill,Skill&$filter=EmployeeID eq '" + itemsRRF["EmployeeID"] + "'");
                                  endpointRequest.Method = "GET";
                                  //if (XML == false)
                                  endpointRequest.Accept = "application/json;odata=verbose";
                                  endpointRequest.UseDefaultCredentials = false;

                                  endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                                HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                                WebResponse webResponse = request.GetResponse();
                                Stream webStream = webResponse.GetResponseStream();
                                StreamReader responseReader = new StreamReader(webStream);
                                string response = responseReader.ReadToEnd();
                                JObject jobj = JObject.Parse(response);
                                jarr = (JArray)jobj["d"]["results"];
                                JArray jarrPT = new JArray();
                                foreach (JObject j in jarr)
                                {
                                    JObject jPT = new JObject();
                                    string RoleBand = j["RoleBand"].ToString();

                                    string EmployeeRole = j["EmployeeRole"].ToString();

                                    string SubPractice = j["SubPractice"].ToString();

                                    string PrimarySkill = j["PrimarySkill"].ToString();

                                    string Skill = j["Skill"].ToString();
                                    string AllocatedResourceName = j["ResourceName"].ToString();

                                    if (variablename == "ReleaseRoleBand")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", RoleBand);
                                    }
                                    else if (variablename == "ReleaseEmploymentRole	")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", EmployeeRole);
                                    }
                                    else if (variablename == "ReleaseSubPractice")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", SubPractice);
                                    }
                                    else if (variablename == "ReleasePrimarySkill")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", PrimarySkill);
                                    }
                                    else if (variablename == "ReleaseSecondarySkill")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", Skill);
                                    }
                                    else if (variablename == "AllocatedResourceName")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", AllocatedResourceName);
                                    }
                                    else
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", "");
                                    }


                                }
                            }
                            else if (Flag == 4)
                            {
                                var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/lists/getbytitle('RMOResourceAssignment')/items?$select=StartDate,EndDate,AllocationPercent,ProjectLoaction&$filter=RRFNumber eq '" + itemsRRF["RRFNO"] + "'");
                                request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                                request.Method = WebRequestMethods.Http.Get;
                                request.Accept = "application/json;odata=verbose";
                                // request.ContentType = "application/json;odata=verbose";
                                request.ContentLength = 0;

                                var securePassword = new SecureString();
                                foreach (char c in Password)
                                {
                                    securePassword.AppendChar(c);
                                }
                                request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);

                                /* HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(URL + "_api/web/lists/getbytitle('RMOResourceAssignment')/items?$select=StartDate,EndDate,AllocationPercent,ProjectLoaction&$filter=RRFNumber eq '" + itemsRRF["RRFNO"] + "'");
                                 endpointRequest.Method = "GET";
                                 //if (XML == false)
                                 endpointRequest.Accept = "application/json;odata=verbose";
                                 endpointRequest.UseDefaultCredentials = false;

                                 endpointRequest.CookieContainer = claimsHelper.CookieContainer; //In case of online*/

                                HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                                WebResponse webResponse = request.GetResponse();
                                Stream webStream = webResponse.GetResponseStream();
                                StreamReader responseReader = new StreamReader(webStream);
                                string response = responseReader.ReadToEnd();
                                JObject jobj = JObject.Parse(response);
                                jarr = (JArray)jobj["d"]["results"];
                                JArray jarrPT = new JArray();
                                foreach (JObject j in jarr)
                                {
                                    JObject jPT = new JObject();
                                    string StartDate = j["StartDate"].ToString();

                                    string EndDate = j["EndDate"].ToString();

                                    string AllocationPer = j["AllocationPercent"].ToString();

                                    string ProjectLocation = j["ProjectLoaction"].ToString();



                                    if (variablename == "AllocationStartDate")
                                    {
                                        if ((itemsRRF[FieldValue].ToString().Contains(" PM") && itemsRRF[FieldValue].ToString().Contains("/") || (itemsRRF[FieldValue].ToString().Contains(" PM") && itemsRRF[FieldValue].ToString().Contains("-"))))
                                        {
                                            StartDate = Convert.ToDateTime(itemsRRF[FieldValue]).ToShortDateString();
                                        }
                                        else if ((itemsRRF[FieldValue].ToString().Contains(" AM") && itemsRRF[FieldValue].ToString().Contains("/") || (itemsRRF[FieldValue].ToString().Contains(" AM") && itemsRRF[FieldValue].ToString().Contains("-"))))
                                        {
                                            StartDate = Convert.ToDateTime(itemsRRF[FieldValue]).ToShortDateString();
                                        }
                                        else
                                        {
                                            StartDate = "";
                                        }
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", StartDate);
                                    }
                                    else if (variablename == "AllocationEndDate")
                                    {
                                        if ((itemsRRF[FieldValue].ToString().Contains(" PM") && itemsRRF[FieldValue].ToString().Contains("/") || (itemsRRF[FieldValue].ToString().Contains(" PM") && itemsRRF[FieldValue].ToString().Contains("-"))))
                                        {
                                            EndDate = Convert.ToDateTime(itemsRRF[FieldValue]).ToShortDateString();
                                        }
                                        else if ((itemsRRF[FieldValue].ToString().Contains(" AM") && itemsRRF[FieldValue].ToString().Contains("/") || (itemsRRF[FieldValue].ToString().Contains(" AM") && itemsRRF[FieldValue].ToString().Contains("-"))))
                                        {
                                            EndDate = Convert.ToDateTime(itemsRRF[FieldValue]).ToShortDateString();
                                        }
                                        else
                                        {
                                            EndDate = "";
                                        }
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", EndDate);
                                    }
                                    else if (variablename == "AllocationPercentage")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", AllocationPer);
                                    }
                                    else if (variablename == "ProjectLocation")
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", ProjectLocation);
                                    }
                                    else
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", "");
                                    }

                                }
                            }
                            else
                            {
                                //var lookFieldvalue = (itemsRRF[FieldValue]) as FieldLookupValue;
                                if (type == "Body")
                                {
                                    if(_EventId == "64" && variablename == "ReleasePendingWithUser")
                                    {
                                        if (FinalTo.ToString().Contains('@'))
                                        {
                                            var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/siteusers?$select=Email,Title&$filter=Email eq " + FinalTo + "");
                                            request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                                            request.Method = WebRequestMethods.Http.Get;
                                            request.Accept = "application/json;odata=verbose";
                                            request.ContentLength = 0;

                                            var securePassword = new SecureString();
                                            foreach (char c in Password)
                                            {
                                                securePassword.AppendChar(c);
                                            }
                                            request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);
                                            HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                                            WebResponse webResponse = request.GetResponse();
                                            Stream webStream = webResponse.GetResponseStream();
                                            StreamReader responseReader = new StreamReader(webStream);
                                            string response = responseReader.ReadToEnd();
                                            JObject jobj = JObject.Parse(response);
                                            jarr = (JArray)jobj["d"]["results"];
                                            JArray jarrPT = new JArray();
                                            foreach (JObject j in jarr)
                                            {
                                                string Vlaue = j["Title"].ToString();
                                                stringBuilder.Replace(variablename + "&#125;", Vlaue);
                                                stringBuilder.Replace("&#123;", "");
                                            }
                                        }
                                        else
                                        {
                                            var request = (HttpWebRequest)WebRequest.Create(URL + "_api/web/siteusers?$select=Email,Title&$filter=Id eq " + FinalTo + "");
                                            request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                                            request.Method = WebRequestMethods.Http.Get;
                                            request.Accept = "application/json;odata=verbose";
                                            request.ContentLength = 0;

                                            var securePassword = new SecureString();
                                            foreach (char c in Password)
                                            {
                                                securePassword.AppendChar(c);
                                            }
                                            request.Credentials = new SharePointOnlineCredentials(UserName, securePassword);
                                            HttpWebResponse endpointResponse = (HttpWebResponse)request.GetResponse();
                                            WebResponse webResponse = request.GetResponse();
                                            Stream webStream = webResponse.GetResponseStream();
                                            StreamReader responseReader = new StreamReader(webStream);
                                            string response = responseReader.ReadToEnd();
                                            JObject jobj = JObject.Parse(response);
                                            jarr = (JArray)jobj["d"]["results"];
                                            JArray jarrPT = new JArray();
                                            foreach (JObject j in jarr)
                                            {
                                                string Vlaue = j["Title"].ToString();                                               
                                                stringBuilder.Replace( variablename + "&#125;", Vlaue);
                                                stringBuilder.Replace("&#123;", "");
                                            }
                                        }

                                    }
                                    else if (itemsRRF[FieldValue] != null)
                                    {
                                        string Vlaue = ((Microsoft.SharePoint.Client.FieldLookupValue)itemsRRF[FieldValue]).LookupValue;
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", Vlaue);
                                    }
                                    else
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", "");
                                    }
                                }
                                else
                                {
                                    if (itemsRRF[FieldValue] != null)
                                    {
                                        string Vlaue = ((Microsoft.SharePoint.Client.FieldLookupValue)itemsRRF[FieldValue]).LookupValue;
                                        stringBuilder.Replace("{" + variablename + "}", Vlaue);
                                    }
                                    else
                                    {
                                        stringBuilder.Replace("&#123;" + variablename + "&#125;", "");
                                    }
                                }

                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                SendMailsDatBaseProfile("pankaj.singh@e2eprojects.com","uday.s@e2eprojects.com",ex.ToString(), "Error ReplaceVariablevalue1");
                 ErrorFlag = "1";
            }
        }

        private static void SendMailsDatBaseProfile(string To, string CC , string Body , string Subject)
        {

            SqlCommand cmdExec = new SqlCommand("exec msdb.dbo.sp_send_dbmail @Profile_name=@Profile_name1," +
                                     "@recipients=@recipients1,@copy_recipients=@copy_recipients1,@subject=@subject1,@body=@body1,@body_format=@body_format1", con);
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            try
            {

                cmdExec.Parameters.AddWithValue("@Profile_name1", "RMO");
                cmdExec.Parameters.AddWithValue("@recipients1", "Geetha.sk@e2eprojects.com");
                cmdExec.Parameters.AddWithValue("@subject1", Subject);
                cmdExec.Parameters.AddWithValue("@body1", Body);
                cmdExec.Parameters.AddWithValue("@copy_recipients1", "Shweta.M@e2eprojects.com");
                //cmdExec.Parameters.AddWithValue("@blind_copy_recipients1", "Shweta.M@e2eprojects.com");

                cmdExec.Parameters.AddWithValue("@body_format1", "HTML");
                cmdExec.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                ex.ToString();
                throw;
            }
        }

    }
}