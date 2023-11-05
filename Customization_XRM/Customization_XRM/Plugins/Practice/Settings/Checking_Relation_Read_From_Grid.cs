using CRM;
using Customization_XRM.Base.PluginBase;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Customization_XRM.Plugins.Practice.Settings
{
    public class Checking_Relation_Read_From_Grid : PluginBase
    {
        public Checking_Relation_Read_From_Grid() : base(typeof(Checking_Relation_Read_From_Grid))
        {

        }

        protected override void ExecutePluginLogic(LocalPluginExecution localPluginExecution)
        {
            if (localPluginExecution == null)
            {
                throw new InvalidPluginExecutionException("Local Plugin Execution is not initialized correctly.");
            }

            //initialize plugin basec components
            IPluginExecutionContext context = localPluginExecution.pluginContext;
            IOrganizationService crmService = localPluginExecution.orgService;
            ITracingService tracingService = localPluginExecution.tracingService;

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                new_Settings settings = crmService.Retrieve(new_Settings.EntityLogicalName, entity.Id, new ColumnSet(true)) as new_Settings;

                var settingsDetailQuery = new QueryExpression
                {
                    EntityName = new_project.EntityLogicalName,
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression
                            {
                                //نام فیلد لوک آپ
                                AttributeName = "new_relatedprojectid",
                                Operator = ConditionOperator.Equal,
                                Values = {entity.Id}
                            }, new ConditionExpression
                            {
                                AttributeName = "statecode",
                                Operator = ConditionOperator.Equal,
                                Values = {0}
                            }
                        }
                    }
                };

                var detailsItem = crmService.RetrieveMultiple(settingsDetailQuery).Entities.Cast<new_project>().ToList();
                foreach (var item in detailsItem)
                {
                    string x = string.Empty;
                    x += item.new_name;
                }
            }
        }
    }
}