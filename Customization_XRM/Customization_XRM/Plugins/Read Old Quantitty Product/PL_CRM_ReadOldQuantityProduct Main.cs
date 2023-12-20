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
    public class PL_CRM_ReadOldQuantityProductMain : PluginBase
    {
        public PL_CRM_ReadOldQuantityProductMain() : base(typeof(PL_CRM_ReadOldQuantityProductMain))
        {

        }

        //EntityReference FLookUpContactTarget;
        //Entity RetriveFLookUpContactTarget;

        EntityReference FLookUpContact;
        Entity RetriveFLookUpContact;       

        EntityReference FLookUpAccount;
        Entity RetriveFLookUpAccount;

        EntityReference FLookUpContactProduct;
        Entity RetriveFLookUpContactProduct;

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
                    var opportunityProduct = crmService.Retrieve(OpportunityProduct.EntityLogicalName, entity.Id, new ColumnSet(true)) as OpportunityProduct;
                    var h = opportunityProduct.ProductId;

                    var FLookUpContactTarget = (EntityReference)(entity.Attributes["productid"]);
                    var RetriveFLookUpContactTarget = crmService.Retrieve(FLookUpContactTarget.LogicalName, FLookUpContactTarget.Id, new ColumnSet("productid"));

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

                            foreach (var i in contactProducts)
                            {
                                FLookUpContactProduct = (EntityReference)(i.Attributes["productid"]);
                                RetriveFLookUpContactProduct = crmService.Retrieve(FLookUpContactProduct.LogicalName, FLookUpContactProduct.Id, new ColumnSet(true));

                                listProductContact.Add(RetriveFLookUpContactProduct);
                            }
                            
                            listProductContact.AddRange(contactProducts);
                        }

                        var x = listProductContact.Where(p => p.GetAttributeValue<int>("new_old_quantity").ToString() != null);

                        var selectContactProductTargenFormList = RetriveFLookUpContactTarget.GetAttributeValue<Guid>("productid").ToString();

                        var contactProductMain = listProductContact.Where(p => p.Attributes["productid"].ToString() == selectContactProductTargenFormList /*&& 
                                                                            p.Attributes["new_old_quantity"] != null*/);
                        var contactMax = contactProductMain.Max(p => p.Attributes["createdon"]);
                        var contactProductMainMax = contactProductMain.Where(p => p.Attributes["createdon"] == contactMax);                       
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

                        var accountProductMain = listProductAccount.Where(p => p.Attributes["productid"] == entity.Attributes["productid"]);
                        var accountProductMainMax = accountProductMain.Max(p => p.Attributes["createdon"]);
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
