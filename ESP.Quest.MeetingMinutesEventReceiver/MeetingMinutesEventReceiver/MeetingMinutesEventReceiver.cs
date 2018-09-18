using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using ESPContractsInfoDataAccess;
using System.Linq;
using Microsoft.SharePoint.Utilities;
using System.Text;
using System.Net.Mail;

namespace ESP.Quest.MeetingMinutesEventReceiver
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class MeetingMinutesEventReceiver : SPItemEventReceiver
    {

        /// <summary>
        /// An item is being added.
        /// </summary>
        /// 
        public override void ItemAdded(SPItemEventProperties properties)
        {

          //utilityMail( "trace", "before try");

            try
            {

               
              
                ContractsInfoEntities1 context = new ContractsInfoEntities1();
                

                //get value of Meeing Type
                string mtgType = properties.ListItem["__x007b_10a0afeb_dd0e_4f1f_9770_9382c0695941_x007d_"].ToString();
               
               
                if (mtgType.Contains("Contract Modification"))
                {
                   // utilityMail("trace", "before projectID");
                    string projectID = properties.ListItem["_941bb36f_c77f_43f6_9925_4259fbffcf33"].ToString();

                    //utilityMail("trace", "before staffing");
                    string staffingChanges = properties.ListItem["_d4db9a9a_c19c_478d_ba16_0bc66639876c"].ToString();
                    //utilityMail("trace", "before cost");
                    string costChanges = properties.ListItem["_27be1220_8724_4fdd_b9f3_5e00f5f72f54"].ToString();
                    //utilityMail("trace", "before schedule");
                    string scheduleChanges = properties.ListItem["_bc91d7ca_ba3b_4c41_acfb_9ed0414f903e"].ToString();
                    //utilityMail("trace", "before scope");
                    string scopeChanges = properties.ListItem["_058687d3_de90_41e8_82ff_959b7e21b098"].ToString();
                    //utilityMail("trace", "before admincomments");
                    string adminComments = properties.ListItem["_b53dbd49_8df9_4382_b3ba_9ea5798e362f"].ToString();
                    
                    //send email when meeting type is Contract Modification
                 
                    //get Project Director, Operations VP and Financial Analyst values from database based on Core Charge #
                    //use Exchange Group for PMO, Pricing and FSO emails


                    var vals = (from c in context.ContractsLists
                                where c.Base_Period_Core_Charge__ == projectID
                                select new
                                {
                                    Director = c.Director_,
                                    VP = c.Vice_President_of_Operations,
                                    FinacialAnalyst = c.Financial_Analyst,
                                    ContractAnalyst = c.Contract_Analyst,
                                    TOTitle = c.Title_of_Task_Order,
                                    Contract = c.Contract_Short_Name,
                                    Prime = c.Prime,
                                    TO = c.Prime_TO_
                                }).FirstOrDefault();

                    string url = properties.Web.Url.Substring(0, properties.Web.Url.IndexOf("/sites"));

                    string directorsEmail = string.Empty;
                    string vpEmail = string.Empty;
                    string financeEmail = string.Empty;
                    string contractAnalystEmail = string.Empty;
                    bool isDirector = false;
                    bool isFinance = false;
                    bool isContract = false;
                    bool isVP = false;
                    using (SPWeb contractSite = new SPSite(url + "/sites/contracts").RootWeb)
                    {
                        SPList contractPOC = contractSite.Lists["Contract POC"];
                        SPListItemCollection items = contractPOC.Items;


                      


                        foreach (SPListItem item in items)
                        {
                            //utilityMail("trace", "before poc name");
                            if (item["Name"].ToString() == vals.Director)
                            {
                               // utilityMail("trace", "before director");
                                directorsEmail = item["E-Mail"].ToString();
                                isDirector = true;
                            }
                            if (item["Name"].ToString() == vals.FinacialAnalyst)
                            {
                                financeEmail = item["E-Mail"].ToString();
                                isFinance = true;
                            }
                            if (item["Name"].ToString() == vals.ContractAnalyst)
                            {
                                contractAnalystEmail = item["E-Mail"].ToString();
                                isContract = true;
                            }
                            if (item["Name"].ToString() == vals.VP)
                            {
                                vpEmail = item["E-Mail"].ToString();
                                isVP = true;
                            }
                        }

                    }

                    //utilityMail("trace", "before start of subject string builder");
                    StringBuilder emailSubject = new StringBuilder("A Contract Modification Announcement for ");
                    emailSubject.Append(vals.Prime + " ");
                    emailSubject.Append(vals.TOTitle + " ");
                    emailSubject.Append("(" + projectID + ") ");
                    emailSubject.Append("has been posted to QUEST ");

                    //create HTML email body
                    //utilityMail("trace", "before start of body stringbuilder");
                    StringBuilder emailBody = new StringBuilder(emailSubject.ToString() + "with the following comments:<br/><br/>");
                    emailBody.Append("<ul>");
                    emailBody.Append("<li><b>Changes to team / subcontractors:</b><br/>");
                    //utilityMail("trace", staffingChanges);
                    emailBody.Append(staffingChanges);
                    emailBody.Append("<br/><br/><br/>");
                    emailBody.Append("</li>");

                    emailBody.Append("<li><b>Changes to contract ceiling value:</b><br/>");
                    //utilityMail("trace", costChanges);
                    emailBody.Append(costChanges);
                    emailBody.Append("<br/><br/><br/>");
                    emailBody.Append("</li>");

                    emailBody.Append("<li><b>Changes to contract schedule / period of performance :</b><br/>");
                     // utilityMail("trace",scheduleChanges );
                    emailBody.Append(scheduleChanges);
                    emailBody.Append("<br/><br/><br/>");
                    emailBody.Append("</li>");

                    emailBody.Append("<li><b>Changes to contract scope:</b><br/>");
                    //utilityMail("trace", scopeChanges);
                    emailBody.Append(scopeChanges);
                    emailBody.Append("<br/><br/>");
                    emailBody.Append("</li>");

                    emailBody.Append("<li><b>Administrative Comments:</b><br/>");
                    //utilityMail("trace", adminComments);
                    emailBody.Append(adminComments);
                    emailBody.Append("<br/><br/>");
                    emailBody.Append("</li>");
                    emailBody.Append("</ul>");
                    emailBody.Append("<br/><br/><br/>");
                    emailBody.Append("<b> Project Director:</b>  Per <b>QWI 101.5 (Manage Contract Modification)</b>, if the contract modification changes the contract ceiling, scope, period of performance, and/or adds or removes a subcontractor, the Project Director will update the spend plan <b>within ten business days after the Contract Modification Announcement.</b>");
                                     
                    emailBody.Append("<br/><br/>");
                    emailBody.Append("<b>Project Management Office:  Per <b>QWI 101.5 (Manage Contract Modification)</b>, if the contract modification changes the period of performance or CDRL schedule, the Project Management Office will update the project CDRL schedule <b>within five business days after the Contract Modification Announcement.</b>");
                    emailBody.Append("<br/><br/>");
                     // utilityMail("trace", url);
                    emailBody.Append("<a href=" + url + "/sites/espquest/meetingminuteforms2>View Item</a>");
                    emailBody.Append("<br/><br/>");
                    emailBody.Append("<center>********CONFIDENTIALITY NOTICE********</center>");
                    emailBody.Append("<br/><br/>");
                    emailBody.Append("This email message, including any attachments, may contain confidential information exclusively provided for intended recipients or authorized representatives of the intended recipients.  Any dissemination of this e-mail by anyone other than an intended recipient or authorized representatives of the intended recipients is strictly prohibited. If you are not a named recipient or authorized representatives of the intended recipients, you are prohibited from any further viewing of the e-mail or any attachments or from making any use of the e-mail or attachments. If you believe you have received this e-mail in error, notify the sender immediately and permanently delete the e-mail, any attachments, and all copies thereof from any drives or storage media and destroy any printouts of the e-mail or attachments.");
                    MailMessage msg = new MailMessage();
                    msg.To.Add("ESPPMO@espus.com");
                    //msg.To.Add("lori.przywozny@espus.com");
                    if(isDirector)
                    msg.To.Add(directorsEmail);
                    if (isVP)                   
                    msg.To.Add(vpEmail);
                    if(isFinance)
                    msg.To.Add(financeEmail);
                    //Removed Pricing per requiremts on 3/10/14
                    //msg.To.Add("pricing@espcorp.org");
                    msg.To.Add("esppmo@espcorp.org");
                    if(isContract)
                    msg.CC.Add(contractAnalystEmail);
                    msg.IsBodyHtml = true;
                    msg.From =  new MailAddress("SharePointAdmin@espcorp.org");
                    msg.Subject = emailSubject.ToString();
                    msg.Body = emailBody.ToString();

                    var client = new SmtpClient("mail2.espcorp.org");
                   
                    //next line is for testing email with local smtp server
                    //var client = new SmtpClient("localhost", 25);
                    client.Send(msg);


                }
            }
            catch (Exception ex)
            {

               // throw new Exception(ex.StackTrace + "####" + ex.Message, ex);

                utilityMail("Error", ex.StackTrace);
               
            }
            
        }

        private static void utilityMail(string subject, string msg)
        {
            MailMessage b = new MailMessage();
            b.To.Add("SharePointAdmin@espus.com");
            b.IsBodyHtml = true;
            b.From = new MailAddress("SharePointAdmin@espus.com");
            b.Subject = subject;
            b.Body = msg;

            var clientb = new SmtpClient("mail2.espcorp.org");
            clientb.Send(b);
        }

        public override void ItemUpdated(SPItemEventProperties properties)
        {

            //utilityMail( "trace", "before try");

            try
            {



                ContractsInfoEntities1 context = new ContractsInfoEntities1();


                //get value of Meeing Type
                string mtgType = properties.ListItem["__x007b_10a0afeb_dd0e_4f1f_9770_9382c0695941_x007d_"].ToString();


                if (mtgType.Contains("Contract Modification"))
                {
                    // utilityMail("trace", "before projectID");
                    string projectID = properties.ListItem["_941bb36f_c77f_43f6_9925_4259fbffcf33"].ToString();

                    //utilityMail("trace", "before staffing");
                    string staffingChanges = properties.ListItem["_d4db9a9a_c19c_478d_ba16_0bc66639876c"].ToString();
                    //utilityMail("trace", "before cost");
                    string costChanges = properties.ListItem["_27be1220_8724_4fdd_b9f3_5e00f5f72f54"].ToString();
                    //utilityMail("trace", "before schedule");
                    string scheduleChanges = properties.ListItem["_bc91d7ca_ba3b_4c41_acfb_9ed0414f903e"].ToString();
                    //utilityMail("trace", "before scope");
                    string scopeChanges = properties.ListItem["_058687d3_de90_41e8_82ff_959b7e21b098"].ToString();
                    //utilityMail("trace", "before admincomments");
                    string adminComments = properties.ListItem["_b53dbd49_8df9_4382_b3ba_9ea5798e362f"].ToString();

                    //send email when meeting type is Contract Modification

                    //get Project Director, Operations VP and Financial Analyst values from database based on Core Charge #
                    //use Exchange Group for PMO, Pricing and FSO emails


                    var vals = (from c in context.ContractsLists
                                where c.Base_Period_Core_Charge__ == projectID
                                select new
                                {
                                    Director = c.Director_,
                                    VP = c.Vice_President_of_Operations,
                                    FinacialAnalyst = c.Financial_Analyst,
                                    ContractAnalyst = c.Contract_Analyst,
                                    TOTitle = c.Title_of_Task_Order,
                                    Contract = c.Contract_Short_Name,
                                    Prime = c.Prime,
                                    TO = c.Prime_TO_
                                }).FirstOrDefault();

                    string url = properties.Web.Url.Substring(0, properties.Web.Url.IndexOf("/sites"));

                    string directorsEmail = string.Empty;
                    string vpEmail = string.Empty;
                    string financeEmail = string.Empty;
                    string contractAnalystEmail = string.Empty;
                    bool isDirector = false;
                    bool isFinance = false;
                    bool isContract = false;
                    bool isVP = false;
                    using (SPWeb contractSite = new SPSite(url + "/sites/contracts").RootWeb)
                    {
                        SPList contractPOC = contractSite.Lists["Contract POC"];
                        SPListItemCollection items = contractPOC.Items;





                        foreach (SPListItem item in items)
                        {
                            //utilityMail("trace", "before poc name");
                            if (item["Name"].ToString() == vals.Director)
                            {
                                // utilityMail("trace", "before director");
                                directorsEmail = item["E-Mail"].ToString();
                                isDirector = true;
                            }
                            if (item["Name"].ToString() == vals.FinacialAnalyst)
                            {
                                financeEmail = item["E-Mail"].ToString();
                                isFinance = true;
                            }
                            if (item["Name"].ToString() == vals.ContractAnalyst)
                            {
                                contractAnalystEmail = item["E-Mail"].ToString();
                                isContract = true;
                            }
                            if (item["Name"].ToString() == vals.VP)
                            {
                                vpEmail = item["E-Mail"].ToString();
                                isVP = true;
                            }
                        }

                    }

                    //utilityMail("trace", "before start of subject string builder");
                    StringBuilder emailSubject = new StringBuilder("A Contract Modification Announcement for ");
                    emailSubject.Append(vals.Prime + " ");
                    emailSubject.Append(vals.TOTitle + " ");
                    emailSubject.Append("(" + projectID + ") ");
                    emailSubject.Append("has been posted to QUEST ");

                    //create HTML email body
                    //utilityMail("trace", "before start of body stringbuilder");
                    StringBuilder emailBody = new StringBuilder(emailSubject.ToString() + "with the following comments:<br/><br/>");
                    emailBody.Append("<ul>");
                    emailBody.Append("<li><b>Changes to team / subcontractors:</b><br/>");
                    //utilityMail("trace", staffingChanges);
                    emailBody.Append(staffingChanges);
                    emailBody.Append("<br/><br/><br/>");
                    emailBody.Append("</li>");

                    emailBody.Append("<li><b>Changes to contract ceiling value:</b><br/>");
                    //utilityMail("trace", costChanges);
                    emailBody.Append(costChanges);
                    emailBody.Append("<br/><br/><br/>");
                    emailBody.Append("</li>");

                    emailBody.Append("<li><b>Changes to contract schedule / period of performance :</b><br/>");
                    // utilityMail("trace",scheduleChanges );
                    emailBody.Append(scheduleChanges);
                    emailBody.Append("<br/><br/><br/>");
                    emailBody.Append("</li>");

                    emailBody.Append("<li><b>Changes to contract scope:</b><br/>");
                    //utilityMail("trace", scopeChanges);
                    emailBody.Append(scopeChanges);
                    emailBody.Append("<br/><br/>");
                    emailBody.Append("</li>");

                    emailBody.Append("<li><b>Administrative Comments:</b><br/>");
                    //utilityMail("trace", adminComments);
                    emailBody.Append(adminComments);
                    emailBody.Append("<br/><br/>");
                    emailBody.Append("</li>");
                    emailBody.Append("</ul>");
                    emailBody.Append("<br/><br/><br/>");
                    emailBody.Append("<b> Project Director:</b>  Per <b>QWI 101.5 (Manage Contract Modification)</b>, if the contract modification changes the contract ceiling, scope, period of performance, and/or adds or removes a subcontractor, the Project Director will update the spend plan <b>within ten business days after the Contract Modification Announcement.</b>");

                    emailBody.Append("<br/><br/>");
                    emailBody.Append("<b>Project Management Office:  Per <b>QWI 101.5 (Manage Contract Modification)</b>, if the contract modification changes the period of performance or CDRL schedule, the Project Management Office will update the project CDRL schedule <b>within five business days after the Contract Modification Announcement.</b>");
                    emailBody.Append("<br/><br/>");
                    // utilityMail("trace", url);
                    emailBody.Append("<a href=" + url + "/sites/espquest/meetingminuteforms2>View Item</a>");
                    emailBody.Append("<br/><br/>");
                    emailBody.Append("<center>********CONFIDENTIALITY NOTICE********</center>");
                    emailBody.Append("<br/><br/>");
                    emailBody.Append("This email message, including any attachments, may contain confidential information exclusively provided for intended recipients or authorized representatives of the intended recipients.  Any dissemination of this e-mail by anyone other than an intended recipient or authorized representatives of the intended recipients is strictly prohibited. If you are not a named recipient or authorized representatives of the intended recipients, you are prohibited from any further viewing of the e-mail or any attachments or from making any use of the e-mail or attachments. If you believe you have received this e-mail in error, notify the sender immediately and permanently delete the e-mail, any attachments, and all copies thereof from any drives or storage media and destroy any printouts of the e-mail or attachments.");
                    MailMessage msg = new MailMessage();
                    msg.To.Add("ESPPMO@espus.com");
                    //msg.To.Add("lori.przywozny@espus.com");
                    if (isDirector)
                        msg.To.Add(directorsEmail);
                    if (isVP)
                        msg.To.Add(vpEmail);
                    if (isFinance)
                        msg.To.Add(financeEmail);
                    //Removed Pricing per requiremts on 3/10/14
                    //msg.To.Add("pricing@espcorp.org");
                    msg.To.Add("esppmo@espcorp.org");
                    if (isContract)
                        msg.CC.Add(contractAnalystEmail);
                    msg.IsBodyHtml = true;
                    msg.From = new MailAddress("SharePointAdmin@espcorp.org");
                    msg.Subject = emailSubject.ToString();
                    msg.Body = emailBody.ToString();

                    var client = new SmtpClient("mail2.espcorp.org");

                    //next line is for testing email with local smtp server
                    //var client = new SmtpClient("localhost", 25);
                    client.Send(msg);


                }
            }
            catch (Exception ex)
            {

                // throw new Exception(ex.StackTrace + "####" + ex.Message, ex);

                utilityMail("Error", ex.StackTrace);

            }

        }
             
        
    }


}