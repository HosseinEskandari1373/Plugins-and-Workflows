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

namespace Customization_XRM
{
    public class CreateNumberAuto : PluginBase
    {
        public CreateNumberAuto() : base(typeof(CreateNumberAuto))
        {

        }

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
                    Entity contact = new Contact();

                    if (entity.LogicalName == contact.LogicalName)
                    {
                        if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                        {
                            //get config table row for empty entity 
                            var fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                              <entity name='contact'>
                                                <attribute name='contactid' />
                                                <attribute name='new_contactnumber' />
                                                <attribute name='createdon' />
                                              </entity>
                                            </fetch>";

                            EntityCollection ecAuto = crmService.RetrieveMultiple(new FetchExpression(fetch));
                            Guid autoNumberRecordId = Guid.Empty;

                            foreach (var itemLookUp in ecAuto.Entities)
                            {
                                autoNumberRecordId = itemLookUp.Id;
                            }

                            //------------------------------------------------

                            //مقدار پیش فرض 
                            var charFix = "C";

                            if (ecAuto.Entities.Count == 0)
                            {
                                var newCounter = Convert.ToInt32("1").ToString().PadLeft(7, '0');
                                entity.Attributes.Add("new_contactnumber", charFix + "-" + newCounter.ToString());
                            }
                            else
                            {
                                //get config table row
                                QueryExpression qe = new QueryExpression("contact");
                                FilterExpression fe = new FilterExpression();
                                qe.ColumnSet = new ColumnSet(true);
                                qe.Orders.Add(new OrderExpression("new_contactnumber", OrderType.Descending));
                                var countContract = crmService.RetrieveMultiple(qe).Entities.First();

                                Entity AutoPost = crmService.Retrieve(entity.LogicalName, countContract.Id, new ColumnSet(true));
                                var currentrecordcounternumber = AutoPost.GetAttributeValue<string>("new_contactnumber");

                                var lencurrentrecordcounternumber = currentrecordcounternumber.Length - 7;
                                var currentrecordcounternumbers = currentrecordcounternumber.Substring(lencurrentrecordcounternumber, 7);

                                //initialize counter 
                                var numbers = string.Concat(currentrecordcounternumbers.Where(char.IsNumber));
                                var lenNum = numbers.Length;

                                var newCounterValue = (Convert.ToInt32(numbers) + 1).ToString().PadLeft(lenNum, '0');

                                entity.Attributes.Add("new_contactnumber", charFix + "-" + newCounterValue.ToString());
                            }
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("The account number can only be set by the system.");
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
