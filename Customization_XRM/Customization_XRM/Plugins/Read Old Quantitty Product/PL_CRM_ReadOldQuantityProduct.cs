using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//namespaces for d365 interaction
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Customization_XRM.Base.PluginBase;
using System.Globalization;
using CRM;

namespace Customization_XRM.Plugins.Read_Old_Quantitty_Product
{
    public class PL_CRM_ReadOldQuantityProduct : PluginBase
    {
        public PL_CRM_ReadOldQuantityProduct() : base(typeof(PL_CRM_ReadOldQuantityProduct))
        {

        }

        EntityReference FLookUpContact;
        Entity RetriveFLookUpContact;       

        EntityReference FLookUpAccount;
        Entity RetriveFLookUpAccount;

        //EntityReference FLookUpContactProduct;
        //Entity RetriveFLookUpContactProduct;

        protected override void ExecutePluginLogic(LocalPluginExecution localPluginExecution)
        {
            try
            {
                if (localPluginExecution == null)
                {
                    throw new InvalidPluginExecutionException("Local Plugin Execution is not initialized correctly.");
                }

                //initialize plugin basec components
                IPluginExecutionContext context = localPluginExecution.pluginContext;
                IOrganizationService crmService = localPluginExecution.orgService;
                ITracingService tracingService = localPluginExecution.tracingService;

                if (context.InputParameters.Contains("Target") &&
                        context.InputParameters["Target"] is Entity)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    //خواندن فرصت مربوط به محصول مورد نیاز 
                    var FLookUpOpurMain = (EntityReference)(entity.Attributes["opportunityid"]);
                    var RetriveFLookUpOpurMain = crmService.Retrieve(FLookUpOpurMain.LogicalName, FLookUpOpurMain.Id, new ColumnSet(true));

                    string contactID = "";
                    string contactName = "";

                    string accountID = "";
                    string accountName = "";

                    // خواندن آیدی شخص مربوط به فرصت
                    if (RetriveFLookUpOpurMain.Attributes["parentcontactid"] != null)
                    {
                        FLookUpContact = (EntityReference)(RetriveFLookUpOpurMain.Attributes["parentcontactid"]);
                        RetriveFLookUpContact = crmService.Retrieve(FLookUpContact.LogicalName, FLookUpContact.Id, new ColumnSet(true));

                        contactID = RetriveFLookUpContact.Attributes["contactid"].ToString();
                        contactName = RetriveFLookUpContact.Attributes["fullname"].ToString();
                    }
                    else
                    {
                        FLookUpAccount = (EntityReference)(RetriveFLookUpOpurMain.Attributes["parentaccountid"]);
                        RetriveFLookUpAccount = crmService.Retrieve(FLookUpAccount.LogicalName, FLookUpAccount.Id, new ColumnSet(true));

                        accountID = RetriveFLookUpAccount.Attributes["accountid"].ToString();
                        accountName = RetriveFLookUpAccount.Attributes["name"].ToString();
                    }

                    if (RetriveFLookUpContact != null)
                    {
                        var contact = crmService.Retrieve(RetriveFLookUpContact.LogicalName, RetriveFLookUpContact.Id, new ColumnSet(true));

                        var contactOpurQuery = new QueryExpression
                        {
                            EntityName = "opportunity",
                            ColumnSet = new ColumnSet(true),
                            Criteria = new FilterExpression
                            {
                                Conditions =
                                {
                                    new ConditionExpression
                                    {
                                        AttributeName = "parentcontactid",
                                        Operator = ConditionOperator.Equal,
                                        Values = {contact.Id}
                                    }
                                }
                            }
                        };

                        var contactOpurs = crmService.RetrieveMultiple(contactOpurQuery).Entities.ToList();
                        List<Entity> listProductContact = new List<Entity>();

                        foreach (var item in contactOpurs)
                        {                           
                            var contactProductQuery = new QueryExpression
                            {
                                EntityName = "opportunityproduct",
                                ColumnSet = new ColumnSet(true),
                                Criteria = new FilterExpression
                                {
                                    Conditions =
                                    {
                                        new ConditionExpression
                                        {
                                            AttributeName = "opportunityid",
                                            Operator = ConditionOperator.Equal,
                                            Values = {item.Id}
                                        }
                                    }
                                }
                            };

                            var contactProducts = crmService.RetrieveMultiple(contactProductQuery).Entities.ToList();

                            listProductContact.AddRange(contactProducts);
                        }
                    }
                    else
                    {
                        var account = crmService.Retrieve(RetriveFLookUpAccount.LogicalName, RetriveFLookUpAccount.Id, new ColumnSet(true));

                        var accountOpurQuery = new QueryExpression
                        {
                            EntityName = "opportunity",
                            ColumnSet = new ColumnSet(true),
                            Criteria = new FilterExpression
                            {
                                Conditions =
                                {
                                    new ConditionExpression
                                    {
                                        AttributeName = "parentcontactid",
                                        Operator = ConditionOperator.Equal,
                                        Values = { account.Id}
                                    }
                                }
                            }
                        };

                        var accountOpurs = crmService.RetrieveMultiple(accountOpurQuery).Entities.ToList();
                        List<Entity> listProductAccount = new List<Entity>();

                        foreach (var item in accountOpurs)
                        {
                            var accountProductQuery = new QueryExpression
                            {
                                EntityName = "opportunityproduct",
                                ColumnSet = new ColumnSet(true),
                                Criteria = new FilterExpression
                                {
                                    Conditions =
                                    {
                                        new ConditionExpression
                                        {
                                            AttributeName = "opportunityid",
                                            Operator = ConditionOperator.Equal,
                                            Values = {item.Id}
                                        }
                                    }
                                }
                            };

                            var accountProducts = crmService.RetrieveMultiple(accountProductQuery).Entities.ToList();

                            listProductAccount.AddRange(accountProducts);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
