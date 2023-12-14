using System;
using System.Linq;

//namespaces for d365 interaction
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Xrm.Sdk.Query;
using Customization_XRM.Base.WorkFlowBase;
using Customization_XRM.Base.Common_URL;

namespace Customization_XRM.Plugins.Practice
{
    public class WF_XRM_ChangeOptionSetDetailForm : WorkflowBase
    {
        public WF_XRM_ChangeOptionSetDetailForm() : base(typeof(WF_XRM_ChangeOptionSetDetailForm))
        {

        }

        [RequiredArgument]
        [Input("Record URL")]
        [ReferenceTarget("")]
        public InArgument<String> RecordURL { get; set; }

        [Input("Child Entity Name")]
        public InArgument<string> ChildEntityName { get; set; }

        [Input("OptionSet Field Name")]
        public InArgument<string> OptionSetFieldName { get; set; }

        [Input("OptionSet Field Value")]
        public InArgument<int> OptionSetFieldValue { get; set; }

        [Input("LookUp Field Name")]
        public InArgument<string> LookUpFieldName { get; set; }

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

            Common objCommon = new Common(localWorkflowExecution.executeContext);
            objCommon.tracingService.Trace("Load CRM Service from context --- OK");

            String _recordURL = this.RecordURL.Get(localWorkflowExecution.executeContext);
            if (_recordURL == null || _recordURL == "")
            {
                return;
            }
            string[] urlParts = _recordURL.Split("?".ToArray());
            string[] urlParams = urlParts[1].Split("&".ToCharArray());
            string ParentObjectTypeCode = urlParams[0].Replace("etc=", "");
            string entityName = objCommon.sGetEntityNameFromCode(ParentObjectTypeCode, objCommon.service);
            string ParentId = urlParams[1].Replace("id=", "");
            Guid entityID = new Guid(ParentId);
            objCommon.tracingService.Trace("ParentObjectTypeCode=" + ParentObjectTypeCode + "--ParentId=" + ParentId);

            Entity entity = crmService.Retrieve(entityName, entityID, new ColumnSet(true));
            var entityItemsQuery = new QueryExpression
            {
                EntityName = ChildEntityName.Get(localWorkflowExecution.executeContext),
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = LookUpFieldName.Get(localWorkflowExecution.executeContext),
                                Operator = ConditionOperator.Equal,
                                Values = { entity.Id}
                            }
                        }
                }
            };

            var entityItems = crmService.RetrieveMultiple(entityItemsQuery).Entities.ToList();

            foreach (var item in entityItems)
            {
                Entity entityToUpdate = new Entity(item.LogicalName, item.Id);
                OptionSetValue OptionVal = new OptionSetValue(OptionSetFieldValue.Get(localWorkflowExecution.executeContext));

                //شده است Set حذف مقدار قبلی که
                entityToUpdate.Attributes.Remove(OptionSetFieldName.Get(localWorkflowExecution.executeContext));
                entityToUpdate.Attributes.Add(OptionSetFieldName.Get(localWorkflowExecution.executeContext), OptionVal);

                crmService.Update(entityToUpdate);
            }
        }
    }
}
