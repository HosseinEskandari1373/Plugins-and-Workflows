using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//namespaces for d365 interaction
using CRM;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Xrm.Sdk.Query;
using System.Globalization;
using Customization_XRM.Base.WorkFlowBase;
using Customization_XRM.Base.Common_URL;

namespace Customization_XRM.Workflows.Return_LookUp_Field
{
    public class ReturnLookUpField : WorkflowBase
    {
        public ReturnLookUpField() : base(typeof(ReturnLookUpField))
        {

        }

        [RequiredArgument]
        [Input("Record URL")]
        [ReferenceTarget("")]
        public InArgument<String> RecordURL { get; set; }

        [Input("First LookUp Field Name")]
        public InArgument<string> FirstLookupFiled { get; set; }

        [Input("Secound LookUp Field Name")]
        public InArgument<string> SecoundLookupField { get; set; }

        [Input("Target Field Name")]
        public InArgument<string> TargetField { get; set; }

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

            Entity entity = crmService.Retrieve(entityName, entityID, new ColumnSet(FirstLookupFiled.Get(localWorkflowExecution.executeContext)));

            //خواندن مقدار لوک آپ فیلد 
            var FLookUp = (EntityReference)(entity.Attributes[FirstLookupFiled.Get(localWorkflowExecution.executeContext)]);
            var RetriveFLookUp = crmService.Retrieve(FLookUp.LogicalName, FLookUp.Id, new ColumnSet(SecoundLookupField.Get(localWorkflowExecution.executeContext)));
            
            var SLookUp = (EntityReference)(RetriveFLookUp.Attributes[SecoundLookupField.Get(localWorkflowExecution.executeContext)]);
            var RetriveSLookUp = crmService.Retrieve(SLookUp.LogicalName, SLookUp.Id, new ColumnSet(TargetField.Get(localWorkflowExecution.executeContext)));

            var TField_Type = RetriveSLookUp.Attributes[TargetField.Get(localWorkflowExecution.executeContext)].GetType();
            var Type = TField_Type.Name;

            if (Type == "OptionSetValue")
            {
                var TField = RetriveSLookUp.GetAttributeValue<OptionSetValue>(TargetField.Get(localWorkflowExecution.executeContext)).Value;
                OutputTargetFieldOptionSet.Set(localWorkflowExecution.executeContext, TField);
            }
            else
            {
                var TFeild1 = (EntityReference)(RetriveSLookUp.Attributes[TargetField.Get(localWorkflowExecution.executeContext)]);
                var RetriveTFeild = crmService.Retrieve(TFeild1.LogicalName, TFeild1.Id, new ColumnSet("name"));
                var nameTargetField = RetriveTFeild["name"].ToString();

                OutputTargetFieldLookUp.Set(localWorkflowExecution.executeContext, nameTargetField);
            }          
        }

        [Output("OutputFieldOptionSetValue")]
        public OutArgument<int> OutputTargetFieldOptionSet { get; set; }

        [Output("OutputFieldLookUpName")]
        public OutArgument<string> OutputTargetFieldLookUp { get; set; }
    }
}
