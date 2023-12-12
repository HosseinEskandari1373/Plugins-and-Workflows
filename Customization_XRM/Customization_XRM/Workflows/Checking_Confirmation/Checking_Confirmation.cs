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
using System.Threading;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace Customization_XRM.Workflows.Checking_Confirmation
{
    public class Checking_Confirmation : WorkflowBase
    {
        public Checking_Confirmation() : base(typeof(Checking_Confirmation))
        {

        }

        [RequiredArgument]
        [Input("Record URL")]
        [ReferenceTarget("")]
        public InArgument<String> RecordURL { get; set; }

        [Input("Confirmed Field Name")]
        public InArgument<string> ConfirmName { get; set; }

        [Input("Sum Minutes Field")]
        public InArgument<string> SumDate { get; set; }

        [Input("Confirm Date Field Name")]
        public InArgument<string> ConfirmDate { get; set; }

        int value;
        int sum;
        bool dateBool;
        DateTime Confirm_Date;
        DateTime CreateOn;
        DateTime LastDate;
        protected override void ExecuteWorkFlowLogic(LocalWorkflowExecution localWorkflowExecution)
        {
            try
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
                objCommon.tracingService.Trace("ParentObjectTypeCode=" + ParentObjectTypeCode + "--ParentId=" + ParentId);

                Guid entityID = new Guid(ParentId);
                Entity entity = crmService.Retrieve(entityName, entityID, new ColumnSet(true));

                QueryExpression confirmQuery = new QueryExpression
                {
                    EntityName = entityName,
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = entityName.ToLower() + "id",
                                Operator = ConditionOperator.Equal,
                                Values = {entity.Id}
                            }
                        }
                    }
                };

                DataCollection<Entity> confirmItemDetails = crmService.RetrieveMultiple(confirmQuery).Entities;

                // Display the results.
                foreach (var item in confirmItemDetails)
                {
                    Confirm_Date = Convert.ToDateTime(item.GetAttributeValue<DateTime>(ConfirmDate.Get(localWorkflowExecution.executeContext)));
                    CreateOn = Convert.ToDateTime(item.GetAttributeValue<DateTime>("createdon"));
                    //LastDate = item.GetAttributeValue<DateTime>("new_last_confirmed_date");
                    dateBool = item.Contains("new_last_confirmed_date");

                    if (dateBool == true)
                    {
                        LastDate = item.GetAttributeValue<DateTime>("new_last_confirmed_date");
                    }
                }

                /*بررسی شرط خالی بودن مقادیر تأیید*/
                if (entity.GetAttributeValue<OptionSetValue>(ConfirmName.Get(localWorkflowExecution.executeContext)) is null)
                {
                    return;
                }

                /// <summary>
                /// مقدار تأیید
                /// </summary>
                /*خواندن مقادیر فیلد تأیید*/
                if (entity.GetAttributeValue<OptionSetValue>(ConfirmName.Get(localWorkflowExecution.executeContext)) != null)
                {
                    value = Convert.ToInt32((entity.GetAttributeValue<OptionSetValue>(ConfirmName.Get(localWorkflowExecution.executeContext)).Value));
                    var attReq_Sale = new RetrieveAttributeRequest()
                    {
                        EntityLogicalName = entityName,
                        LogicalName = ConfirmName.Get(localWorkflowExecution.executeContext),
                        RetrieveAsIfPublished = true
                    };
                    var attResponse_Sale = (RetrieveAttributeResponse)crmService.Execute(attReq_Sale);
                    var attMetadata_Sale = (EnumAttributeMetadata)attResponse_Sale.AttributeMetadata;
                    var sale_Confirm_Label = attMetadata_Sale.OptionSet.Options.Where(x => x.Value == value).FirstOrDefault().Label.UserLocalizedLabel.Label;

                    /*خواندن مقدار فیلد مجموع زمان*/
                    sum = Convert.ToInt32(entity.GetAttributeValue<int>(SumDate.Get(localWorkflowExecution.executeContext)));
                }

                /*بررسی شروط*/
                /*if (entity.GetAttributeValue<OptionSetValue>(ConfirmName.Get(localWorkflowExecution.executeContext)) != null && 
                        (value == 100000000 || value == 100000001 || value == 100000002))*/
                if(value != 0)
                {
                    if (dateBool == false)
                    {
                        //خواندن تاریخ 
                        int sale_Diff = Convert.ToInt32((Confirm_Date - CreateOn).TotalMinutes);
                        int res = (sale_Diff + sum);

                        entity.Attributes.Add(SumDate.Get(localWorkflowExecution.executeContext), res);
                        crmService.Update(entity);

                        //entity.Attributes.Add("new_last_confirmed_date", CreateOn);
                        crmService.Update(entity);
                    }
                    else
                    {
                        //خواندن تاریخ 
                        int sale_Diff = Convert.ToInt32((Confirm_Date - LastDate).TotalMinutes);
                        int res = (sale_Diff + sum);

                        entity.Attributes[SumDate.Get(localWorkflowExecution.executeContext)] = res;
                        //entity.Attributes["new_last_confirmed_date"] = Confirm_Date;

                        crmService.Update(entity);
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
