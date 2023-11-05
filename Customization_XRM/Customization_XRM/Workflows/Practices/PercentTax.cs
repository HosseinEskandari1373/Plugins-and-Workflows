using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

//namespaces for d365 interaction
using CRM;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Xrm.Sdk.Query;
using System.Globalization;
using System.Linq;
using Customization_XRM.Base.WorkFlowBase;

namespace Customization_XRM
{
    public class PercentTax : WorkflowBase
    {
        public PercentTax() : base(typeof(PercentTax))
        {

        }

        [Input("Invoice Id")]
        [ReferenceTarget("invoice")]
        public InArgument<EntityReference> InvoiceId { get; set; }

        [Input("Tax Percent")]
        public InArgument<decimal> Percent { get; set; }

        protected override void ExecuteWorkFlowLogic(LocalWorkflowExecution localWorkflowExecution)
        {
            if (localWorkflowExecution == null)
            {
                throw new InvalidPluginExecutionException("Local Plugin Execution is not initialized correctly.");
            }

            //initialize plugin basec components
            IWorkflowContext context = localWorkflowExecution.pluginContext;
            IOrganizationService crmService = localWorkflowExecution.orgService;
            ITracingService tracingService = localWorkflowExecution.tracingService;

            var invoice = crmService.Retrieve(Invoice.EntityLogicalName, InvoiceId.Get(localWorkflowExecution.executeContext).Id, new ColumnSet(true));
            var invoiceItemsQuery = new QueryExpression
            {
                EntityName = InvoiceDetail.EntityLogicalName,
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "invoiceid",
                                Operator = ConditionOperator.Equal,
                                Values = {invoice.Id}
                            }
                        }
                }
            };
            var invoiceItems = crmService.RetrieveMultiple(invoiceItemsQuery).Entities.Cast<InvoiceDetail>().ToList();
            string des = "";
            bool IsCatch = false;

            foreach (var item in invoiceItems)
            {
                try
                {
                    var x = Percent.Get(localWorkflowExecution.executeContext);
                    crmService.Update(new InvoiceDetail { Id = item.Id, Tax = new Money(item.BaseAmount.Value * Percent.Get(localWorkflowExecution.executeContext)) });
                }
                catch (Exception)
                {
                    IsCatch = true;
                    des = $"محاسبه مالیات برای {item.ProductId.Name} انجام نگردید.";
                }

            }
            if (!IsCatch)
            {
                des = $"محاسبه مالیات در تاریخ {DateTime.Now.ToString("yyyy/MM/dd HH", CultureInfo.GetCultureInfo("fa-ir", "fa"))} با موفقیت انجام گردید.";
            }
            Des.Set(localWorkflowExecution.executeContext, des);
        }

        [Output("Description")]
        public OutArgument<string> Des { get; set; }
    }
}