//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

////namespaces for d365 interaction
//using CRM;
//using Microsoft.Xrm.Sdk;
//using Microsoft.Xrm.Sdk.Workflow;
//using System.Activities;
//using Microsoft.Xrm.Sdk.Query;
//using System.Globalization;
//using Customization_XRM.Base.WorkFlowBase;
//using Customization_XRM.Base.Common_URL;
//using System.Threading;

//namespace Customization_XRM
//{
//    public class CreateNumber : WorkflowBase
//    {
//        public CreateNumber() : base(typeof(CreateNumber))
//        {

//        }

//        [RequiredArgument]
//        [Input("Record URL")]
//        [ReferenceTarget("")]
//        public InArgument<String> RecordURL { get; set; }

//        [Input("Attribute Name")]
//        public InArgument<string> AttributeName { get; set; }

//        [Input("Input Name")]
//        public InArgument<string> InputName { get; set; }

//        [Input("Input Length")]
//        public InArgument<int> InputLength { get; set; }

//        protected override void ExecuteWorkFlowLogic(LocalWorkflowExecution localWorkflowExecution)
//        {
//            try
//            {
//                if (localWorkflowExecution == null)
//                {
//                    throw new InvalidPluginExecutionException("Local Plugin Execution is not initialized correctly.");
//                }

//                //initialize plugin basec components
//                IWorkflowContext context = localWorkflowExecution.pluginContext;
//                IOrganizationService crmService = localWorkflowExecution.orgService;
//                ITracingService tracingService = localWorkflowExecution.tracingService;

//                Common objCommon = new Common(localWorkflowExecution.executeContext);
//                objCommon.tracingService.Trace("Load CRM Service from context --- OK");

//                String _recordURL = this.RecordURL.Get(localWorkflowExecution.executeContext);
//                if (_recordURL == null || _recordURL == "")
//                {
//                    return;
//                }
//                string[] urlParts = _recordURL.Split("?".ToArray());
//                string[] urlParams = urlParts[1].Split("&".ToCharArray());
//                string ParentObjectTypeCode = urlParams[0].Replace("etc=", "");
//                string entityName = objCommon.sGetEntityNameFromCode(ParentObjectTypeCode, objCommon.service);
//                string ParentId = urlParams[1].Replace("id=", "");
//                objCommon.tracingService.Trace("ParentObjectTypeCode=" + ParentObjectTypeCode + "--ParentId=" + ParentId);

//                Guid entityID = new Guid(ParentId);
//                Entity entity = crmService.Retrieve(entityName, entityID, new ColumnSet(true));

//                //get config table row
//                QueryExpression qe = new QueryExpression(entityName);
//                FilterExpression fe = new FilterExpression();
//                qe.ColumnSet = new ColumnSet(true);
//                qe.Orders.Add(new OrderExpression(AttributeName.Get(localWorkflowExecution.executeContext), OrderType.Descending));
//                var countContract = crmService.RetrieveMultiple(qe).Entities.First();

//                Entity AutoPost = crmService.Retrieve(entity.LogicalName, countContract.Id, new ColumnSet(true));
//                var currentrecordcounternumber = AutoPost.GetAttributeValue<string>(AttributeName.Get(localWorkflowExecution.executeContext));

//                //------------------------------------------------
//                //مقدار پیش فرض 
//                var charFix = InputName.Get(localWorkflowExecution.executeContext);

//                if (currentrecordcounternumber == null)
//                {
//                    var newCounter = Convert.ToInt32("1").ToString().PadLeft((InputLength.Get(localWorkflowExecution.executeContext)), '0');

//                    entity.Attributes.Add(AttributeName.Get(localWorkflowExecution.executeContext), charFix + "-" + newCounter);
//                    crmService.Update(entity);
//                }
//                else
//                {
//                    var lencurrentrecordcounternumber = currentrecordcounternumber.Length - InputLength.Get(localWorkflowExecution.executeContext);
//                    var currentrecordcounternumbers = currentrecordcounternumber.Substring(lencurrentrecordcounternumber, InputLength.Get(localWorkflowExecution.executeContext));

//                    //initialize counter 
//                    var numbers = string.Concat(currentrecordcounternumbers.Where(char.IsNumber));
//                    var lenNum = numbers.Length;

//                    var newCounterValue = (Convert.ToInt32(numbers) + 1).ToString().PadLeft(lenNum, '0');

//                    entity.Attributes.Add(AttributeName.Get(localWorkflowExecution.executeContext), charFix + "-" + newCounterValue.ToString());
//                    crmService.Update(entity);
//                }
//            }
//            catch (Exception ex)
//            {
//                throw new InvalidPluginExecutionException(ex.Message);
//            }
//        }
//    }
//}
