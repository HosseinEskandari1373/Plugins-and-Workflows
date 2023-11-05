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

//namespace Customization_XRM.Workflows.Practices
//{
//    public class CreateNumberAuto02 : WorkflowBase
//    {
//        public CreateNumberAuto02() : base(typeof(CreateNumberAuto02))
//        {

//        }

//        [RequiredArgument]
//        [Input("Record URL")]
//        [ReferenceTarget("")]
//        public InArgument<String> RecordURL { get; set; }

//        [Input("Setting ID")]
//        [ReferenceTarget("new_report_daily_performance")]
//        //[ReferenceTarget("new_settings")]
//        public InArgument<EntityReference> SettingID { get; set; }

//        [Input("Attribute Name")]
//        public InArgument<string> AttributeName { get; set; }

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
//                Entity settings = crmService.Retrieve(new_Settings.EntityLogicalName, SettingID.Get(localWorkflowExecution.executeContext).Id, new ColumnSet(true));

//                //Get Value From Settings
//                var preCode = settings.Attributes["new_pre_code"];
//                var lengthCode = settings.Attributes["new_length_code"];
//                var code = settings.Attributes["new_code"];
//                var separator = settings.Attributes["new_separator"];

//                //var maxCode = settings.Attributes["new_max_code"];
//                Object maxCode;

//                if (!settings.Attributes.Contains("new_max_code"))
//                {
//                    //get config table row
//                    QueryExpression qe = new QueryExpression(entityName);
//                    FilterExpression fe = new FilterExpression();
//                    qe.ColumnSet = new ColumnSet(AttributeName.Get(localWorkflowExecution.executeContext));
//                    qe.Orders.Add(new OrderExpression(AttributeName.Get(localWorkflowExecution.executeContext), OrderType.Descending));
//                    var countContract = crmService.RetrieveMultiple(qe).Entities.First();

//                    Entity AutoPost = crmService.Retrieve(entity.LogicalName, countContract.Id, new ColumnSet(true));
//                    var currentrecordcounternumber = AutoPost.GetAttributeValue<string>(AttributeName.Get(localWorkflowExecution.executeContext));

//                    if (currentrecordcounternumber != null)
//                    {
//                        var lencurrentrecordcounternumber = currentrecordcounternumber.Length - Convert.ToInt32(lengthCode);
//                        var currentrecordcounternumbers = currentrecordcounternumber.Substring(lencurrentrecordcounternumber, Convert.ToInt32(lengthCode));

//                        //initialize counter 
//                        var numbers = string.Concat(currentrecordcounternumbers.Where(char.IsNumber));
//                        var lenNum = numbers.Length;

//                        var newCounterValue = (Convert.ToInt32(numbers) + 1).ToString().PadLeft(lenNum, '0');

//                        entity.Attributes.Add(AttributeName.Get(localWorkflowExecution.executeContext), preCode.ToString() + separator.ToString() + newCounterValue.ToString());
//                        crmService.Update(entity);

//                        settings.Attributes.Add("new_max_code", preCode.ToString() + separator.ToString() + newCounterValue.ToString());
//                        crmService.Update(settings);
//                    }
//                    else
//                    {
//                        var newCounter = Convert.ToInt32(code).ToString().PadLeft(Convert.ToInt32(lengthCode), '0');

//                        entity.Attributes.Add(AttributeName.Get(localWorkflowExecution.executeContext), preCode.ToString() + separator.ToString() + newCounter);
//                        crmService.Update(entity);

//                        settings.Attributes.Add("new_max_code", preCode.ToString() + separator.ToString() + newCounter);
//                        crmService.Update(settings);
//                    }
//                }
//                else
//                {
//                    maxCode = settings.Attributes["new_max_code"];
//                    var numbers = string.Concat(maxCode.ToString().Where(char.IsNumber));
//                    var newCounterValue1 = (Convert.ToInt32(numbers) + 1).ToString().PadLeft(Convert.ToInt32(lengthCode), '0');

//                    settings.Attributes.Remove("new_max_code");
//                    crmService.Update(settings);
//                    settings.Attributes.Add("new_max_code", preCode.ToString() + separator + newCounterValue1.ToString());
//                    crmService.Update(settings);

//                    //entity.Attributes.Add(AttributeName.Get(localWorkflowExecution.executeContext), preCode.ToString() + separator + newCounterValue1.ToString());
//                    entity.Attributes.Add(AttributeName.Get(localWorkflowExecution.executeContext), settings.Attributes["new_max_code"]);
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
