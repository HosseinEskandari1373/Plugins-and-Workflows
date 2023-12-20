using System;
using System.Collections.Generic;
using System.Linq;

//namespaces for d365 interaction
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Customization_XRM.Base.PluginBase;
using CRM;

namespace Customization_XRM.Plugins.Read_Old_Quantitty_Product
{
    public class PL_CRM_ReadOldQuantityProduct : PluginBase
    {
        public PL_CRM_ReadOldQuantityProduct() : base(typeof(PL_CRM_ReadOldQuantityProduct))
        {

        }

        EntityReference FLookUpContactTarget;

        EntityReference FLookUpContact;
        Contact contact;

        EntityReference FLookUpAccount;
        Account account;


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
                    OpportunityProduct opportunityProduct = crmService.Retrieve(OpportunityProduct.EntityLogicalName, entity.Id, new ColumnSet(true)) as OpportunityProduct;
                    FLookUpContactTarget = opportunityProduct.ProductId;

                    //خواندن فرصت مربوط به محصول مورد نیاز 
                    var lookUpOpurMain = opportunityProduct.OpportunityId;
                    Opportunity opportunity = crmService.Retrieve(lookUpOpurMain.LogicalName, lookUpOpurMain.Id, new ColumnSet(true)) as Opportunity;

                    Guid contactID;
                    string contactName = "";

                    Guid accountID;
                    string accountName = "";


                    // خواندن آیدی شخص مربوط به فرصت
                    if (opportunity.ParentContactId != null)
                    {
                        FLookUpContact = opportunity.ParentContactId;
                        contact = crmService.Retrieve(FLookUpContact.LogicalName, FLookUpContact.Id, new ColumnSet(true)) as Contact;

                        contactID = contact.Id;
                        contactName = contact.FullName;
                    }
                    else
                    {
                        FLookUpAccount = opportunity.ParentAccountId;
                        account = crmService.Retrieve(FLookUpAccount.LogicalName, FLookUpAccount.Id, new ColumnSet(true)) as Account;

                        accountID = account.Id;
                        accountName = account.Name;
                    }

                    if (contact != null)
                    {
                        Contact readContact = crmService.Retrieve(contact.LogicalName, contact.Id, new ColumnSet(true)) as Contact;

                        QueryExpression contactOpurQuery = new QueryExpression
                        {
                            EntityName = Opportunity.EntityLogicalName,
                            ColumnSet = new ColumnSet(true),
                            Criteria = new FilterExpression
                            {
                                Conditions =
                                {
                                    new ConditionExpression
                                    {
                                        AttributeName = "parentcontactid",
                                        Operator = ConditionOperator.Equal,
                                        Values = { readContact.Id}
                                    }
                                }
                            }
                        };

                        IEnumerable<Opportunity> contactOpurs = crmService.RetrieveMultiple(contactOpurQuery).Entities.Cast<Opportunity>().ToList();
                        List<OpportunityProduct> listProductContact = new List<OpportunityProduct>();

                        foreach (var item in contactOpurs)
                        {
                            QueryExpression contactProductQuery = new QueryExpression
                            {
                                EntityName = OpportunityProduct.EntityLogicalName,
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

                            List<OpportunityProduct> contactProducts = crmService.RetrieveMultiple(contactProductQuery).Entities.Cast<OpportunityProduct>().ToList();
                            IEnumerable<OpportunityProduct> selectNewOldQuantity = contactProducts.Where(p => p.new_old_quantity != null && 
                                                                                                            p.ProductId.Id == FLookUpContactTarget.Id);

                            if (selectNewOldQuantity.Any())
                            {
                                listProductContact.AddRange(selectNewOldQuantity);
                            }
                        }

                        var checkValue = listProductContact.Where(p => p.ProductId.Id == FLookUpContactTarget.Id).Select(p => p.new_old_quantity);
                        int newOldQuantity;

                        if (checkValue.Any())
                        {
                            IEnumerable<OpportunityProduct> selectedTargetProductFromList = listProductContact.Where(p => p.ProductId.Id == FLookUpContactTarget.Id);
                            DateTime? maxCreatedOn = selectedTargetProductFromList.Max(p => p.CreatedOn);

                            IEnumerable<OpportunityProduct> selectedResulrProduct = selectedTargetProductFromList.Where(p => p.CreatedOn == maxCreatedOn);
                            OpportunityProduct selectOneRecord = selectedResulrProduct.FirstOrDefault();
                            newOldQuantity = Convert.ToInt32(selectOneRecord.new_old_quantity);
                        }
                        else
                        {
                            newOldQuantity = 0;
                        }

                        crmService.Update(new OpportunityProduct { Id = entity.Id, new_old_quantity = newOldQuantity });
                    }
                    else
                    {
                        Account readAccount = crmService.Retrieve(account.LogicalName, account.Id, new ColumnSet(true)) as Account;

                        QueryExpression accountOpurQuery = new QueryExpression
                        {
                            EntityName = Opportunity.EntityLogicalName,
                            ColumnSet = new ColumnSet(true),
                            Criteria = new FilterExpression
                            {
                                Conditions =
                                {
                                    new ConditionExpression
                                    {
                                        AttributeName = "parentcontactid",
                                        Operator = ConditionOperator.Equal,
                                        Values = { readAccount.Id}
                                    }
                                }
                            }
                        };

                        IEnumerable<Opportunity> accountOpurs = crmService.RetrieveMultiple(accountOpurQuery).Entities.Cast<Opportunity>().ToList();
                        List<OpportunityProduct> listProductAccount = new List<OpportunityProduct>();

                        foreach (var item in accountOpurs)
                        {
                            QueryExpression accountProductQuery = new QueryExpression
                            {
                                EntityName = OpportunityProduct.EntityLogicalName,
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

                            List<OpportunityProduct> accountProducts = crmService.RetrieveMultiple(accountProductQuery).Entities.Cast<OpportunityProduct>().ToList();
                            IEnumerable<OpportunityProduct> selectNewOldQuantity = accountProducts.Where(p => p.new_old_quantity != null &&
                                                                                                            p.ProductId.Id == FLookUpContactTarget.Id);

                            if (selectNewOldQuantity.Any())
                            {
                                listProductAccount.AddRange(selectNewOldQuantity);
                            }
                        }

                        var checkValue = listProductAccount.Where(p => p.ProductId.Id == FLookUpContactTarget.Id).Select(p => p.new_old_quantity);
                        int newOldQuantity;

                        if (checkValue.Any())
                        {
                            IEnumerable<OpportunityProduct> selectedTargetProductFromList = listProductAccount.Where(p => p.ProductId.Id == FLookUpContactTarget.Id);
                            DateTime? maxCreatedOn = selectedTargetProductFromList.Max(p => p.CreatedOn);

                            IEnumerable<OpportunityProduct> selectedResulrProduct = selectedTargetProductFromList.Where(p => p.CreatedOn == maxCreatedOn);
                            OpportunityProduct selectOneRecord = selectedResulrProduct.FirstOrDefault();
                            newOldQuantity = Convert.ToInt32(selectOneRecord.new_old_quantity);
                        }
                        else
                        {
                            newOldQuantity = 0;
                        }

                        crmService.Update(new OpportunityProduct { Id = entity.Id, new_old_quantity = newOldQuantity });
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
