using System;
using P21.Extensions.BusinessRule;
using System.Net.Mail;
using System.Collections.Generic;
using System.Linq;
using System.Data.SqlClient;





namespace jlm_emailValidation_v02
{
    /*
     * Description: This is a single row rule that will validate that each field passed in has a proper Email format.
     */
    public class jlm_emailValidation_v02 : Rule
    {
        public override RuleResult Execute()
        {
            RuleResult rr = new RuleResult();

            // Fetch Email Address Entered
            string emailToValidate = Data.Fields.GetFieldByAlias("validation_email").FieldValue;

            // If the Email Address is Blank Return True
            if (emailToValidate == "")
            {
                rr.Success = true;
                return rr;
            }

            // Else, Validate
            else
            {
                // Create a DataTable of Restricted Emails to Validate Against
                    string P21ConnectionString = "Data Source=" + Session.Server + ";Initial Catalog=" + Session.Database + ";Integrated Security=SSPI;Application Name=Email_Validation_DynaChange;";
                    string query = "select top 1 c.email_address as 'restricted_email' from p21_view_contacts c(nolock)  where c.email_address is not null";
                    System.Data.DataTable restrictedEmailTable = new System.Data.DataTable();

                try
                {
                    using (SqlConnection connection = new SqlConnection(P21ConnectionString))
                    {
                        connection.Open();
                        using (SqlConnection yourConnection = new SqlConnection(P21ConnectionString))
                        {
                            using (SqlCommand cmd = new SqlCommand(query, yourConnection))
                            {
                                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                {
                                    da.Fill(restrictedEmailTable);
                                }
                            }
                        }
                    }
                }

                catch (Exception ex)
                {
                    restrictedEmailTable.Columns.Add("restrictedEmailTable");
                    rr.Message = ex.Message;
                }
                Console.WriteLine(restrictedEmailTable);
               
                // Split Entered Email Address on ; into an Array
                string[] arrEmailAddress = emailToValidate.Split(';');

                // If the Length of the Array is 1, Validate the Single Email Address
                if (arrEmailAddress.Length == 1)
                {
                    try
                    {
                        var emailValidation = new MailAddress(emailToValidate);

                        // If the Email Address is on the Restricted Email List, Notify and Return False
                        if (restrictedEmailTable.AsEnumerable().Where(c => c.Field<string>("restricted_email").Equals("your lookup value")).Count() > 0;)
                        {
                            try
                            {
                                rr.Message = "Provided Email Address Is Invalid Due to Being a Restricted Email Address. Contact iSupport for Further Information.";
                            }
                            catch
                            {
                                rr.Message = "Provided Email Address Is Invalid Due to Being a Restricted Email Address. Contact iSupport for Further Information.";
                            }
                            rr.Success = false;
                            return rr;
                        }

                        // If the Email Address Entered is an Internal Hydradyne Email Address, Update to noreply@hydradynellc.com
                        else if ((emailValidation.Host == "hydradynellc.com" || emailValidation.Host == "hydra-dyne.com") & emailValidation.Address != "noreply@hydradynellc.com")
                        {
                            Data.Fields.GetFieldByAlias("validation_email").FieldValue = "noreply@hydradynellc.com";
                            rr.Message = "Internal Hydradyne Email Address Changed to 'noreply@hydradynellc.com'";
                            rr.Success = true;
                            return rr;
                        }

                        // Else, the Email Passes Validation
                        else
                        {
                            Data.Fields.GetFieldByAlias("validation_email").FieldValue = emailValidation.Address;
                            rr.Success = true;
                            return rr;
                        }
                    }
                    catch
                    {
                        rr.Message = "An Invalid Email Address was Entered";
                        rr.Success = false;
                        return rr;
                    }
                }

                // If the Length of the Array is GREATER Than 1, Proceed to Loop Over the Array
                else
                {
                    List<string> emailList = new List<string>(); // List of Validated Email Addresses
                    List<string> hydradyneEmails = new List<string>(); // List of Internal Hydradyne Email Addresses Entered
                    List<string> restrictedEmails = new List<string>(); // List of Restricted Emails Addresses Entered
                    List<string> messageUpdates = new List<string>(); // List of Message Prompts

                    foreach (string i in arrEmailAddress)
                    {
                        try
                        {
                            var emailValidation = new MailAddress(i);

                            // If the Email Address is on the Resticted Email List, Set Flag to Notify and Don't Add to Valid List
                            if (restricted_emails.Contains(emailValidation.Address))
                            {
                                restrictedEmails.Add(emailValidation.Address);
                                messageUpdates.Add("Email Address '" + i + "' is a Restricted Email Address and Has Been Removed");
                            }

                            // If the Email Address Entered is an Internal Hydradyne Email Address, Add noreply@hydradynellc.com to List
                            else if ((emailValidation.Host == "hydradynellc.com" || emailValidation.Host == "hydra-dyne.com"))
                            {
                                // If the Email Address != noreply@hydradynellc.com, Set Flag to Notify the Email was Updated
                                if (!hydradyneEmails.Any())
                                {
                                    emailList.Add("noreply@hydradynellc.com");
                                }

                                if (emailValidation.Address != "noreply@hydradynellc.com")
                                {
                                    hydradyneEmails.Add(emailValidation.Address);
                                    messageUpdates.Add("Email Address '" + i + "' is an Internal Address and Has Been Removed");
                                }
                            }

                            // Add Validated Email Address
                            else
                            {
                                emailList.Add(emailValidation.Address);
                                if (i.Equals(emailValidation.Address, StringComparison.OrdinalIgnoreCase) == false)
                                {
                                    messageUpdates.Add("Email Address '" + i + "' Has Been Updated to '" + emailValidation.Address + "'");
                                }
                            }
                        }
                        // Continue if a Single Email Address Fails Validation
                        catch
                        {
                            continue;
                        }
                    }

                    // If the List is Not Empty, Return Validated Email Addresses
                    if (emailList.Count >= 1)
                    {
                        // Remove Duplicates
                        emailList = emailList.Distinct().ToList();

                        // If the Valid Email List is GREATER Than 1, Remove noreply@hydradynellc.com
                        if (emailList.Count > 1 & emailList.Contains("noreply@hydradynellc.com"))
                        {
                            emailList.Remove("noreply@hydradynellc.com");
                        }

                        // Join Remaining Valid Email Addresses Into A String to be Returned
                        var validEmails = String.Join(";", emailList.ToArray());
                        Data.Fields.GetFieldByAlias("validation_email").FieldValue = validEmails;

                        if (messageUpdates.Any())
                        {
                            rr.Message = String.Join(Environment.NewLine, messageUpdates.ToArray());
                        }

                        rr.Success = true;
                        return rr;
                    }

                    // If the List is Empty, All Email Addresses Failed Validation
                    else
                    {
                        rr.Message = "An Invalid Email Address was Entered";
                        rr.Success = false;
                        return rr;
                    }
                }
            }
        }

        public override string GetDescription()
        {
            return "Internally Created DynaChange Business Rule to Validate Email Addresses";
        }

        public override string GetName()
        {
            return "jlm_emailValidation_v02";
        }
    }
}
