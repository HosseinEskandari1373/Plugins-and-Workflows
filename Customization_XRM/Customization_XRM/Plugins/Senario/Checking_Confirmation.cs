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
using Customization_XRM.Base.PluginBase;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Graph;

namespace Customization_XRM.Plugins.Senario
{
    public class Checking_Confirmation : PluginBase
    {

        int saleValue;
        int sum_Sale;

        int sale_engineeringValue;
        int sum_SaleEngineering;

        int serviceValue;
        int sum_Service;
        public Checking_Confirmation() : base(typeof(Checking_Confirmation))
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

            //Target is present or not
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                if (context.MessageName.ToLower() != "update")
                {
                    return;
                }

                Entity entity = context.InputParameters["Target"] as Entity;

                var column = new ColumnSet(true);
                var contract = crmService.Retrieve(new_contract.EntityLogicalName, entity.Id, column) as new_contract;

                /*بررسی شرط خالی بودن مقادیر تأیید*/
                if ((OptionSetValue)contract.new_sale_confirmed is null &&
                        (OptionSetValue)contract.new_sale_engineering_confirmed is null &&
                            (OptionSetValue)contract.new_service_confirmed is null)
                {
                    return;
                }

                /// <summary>
                /// فروش
                /// </summary>
                /*خواندن مقادیر فیلد تأیید فروش*/
                if ((OptionSetValue)contract.new_sale_confirmed != null)
                {
                    saleValue = Convert.ToInt32(((OptionSetValue)contract.new_sale_confirmed).Value);
                    var attReq_Sale = new RetrieveAttributeRequest()
                    {
                        EntityLogicalName = new_contract.EntityLogicalName,
                        LogicalName = "new_sale_confirmed",
                        RetrieveAsIfPublished = true
                    };
                    var attResponse_Sale = (RetrieveAttributeResponse)crmService.Execute(attReq_Sale);
                    var attMetadata_Sale = (EnumAttributeMetadata)attResponse_Sale.AttributeMetadata;
                    var sale_Confirm_Label = attMetadata_Sale.OptionSet.Options.Where(x => x.Value == saleValue).FirstOrDefault().Label.UserLocalizedLabel.Label;

                    /*خواندن مقدار فیلد مجموع زمان فروش*/
                    sum_Sale = Convert.ToInt32(contract.new_sum_sale_time);
                }

                if ((OptionSetValue)contract.new_sale_engineering_confirmed != null)
                {
                    /// <summary>
                    /// مهندسی فروش
                    /// </summary>
                    /*خواندن مقادیر فیلد تأیید مهندسی فروش*/
                    sale_engineeringValue = Convert.ToInt32(((OptionSetValue)contract.new_sale_engineering_confirmed).Value);
                    var attReq_SaleEngineering = new RetrieveAttributeRequest()
                    {
                        EntityLogicalName = new_contract.EntityLogicalName,
                        LogicalName = "new_sale_engineering_confirmed",
                        RetrieveAsIfPublished = true
                    };
                    var attResponse_SaleEngineering = (RetrieveAttributeResponse)crmService.Execute(attReq_SaleEngineering);
                    var attMetadata_SaleEngineering = (EnumAttributeMetadata)attResponse_SaleEngineering.AttributeMetadata;
                    var saleEngineering_Confirm_Label = attMetadata_SaleEngineering.OptionSet.Options.Where(x => x.Value == sale_engineeringValue).FirstOrDefault().Label.UserLocalizedLabel.Label;

                    /*خواندن مقدار فیلد مجموع زمان مهندسی فروش*/
                    sum_SaleEngineering = Convert.ToInt32(contract.new_sum_sale_engineering_time);
                }

                if ((OptionSetValue)contract.new_service_confirmed != null)
                {
                    /// <summary>
                    /// خدمات
                    /// </summary>
                    /*خواندن مقادیر فیلد تأیید خدمات*/
                    serviceValue = Convert.ToInt32(((OptionSetValue)contract.new_service_confirmed).Value);
                    var attReq_Service = new RetrieveAttributeRequest()
                    {
                        EntityLogicalName = new_contract.EntityLogicalName,
                        LogicalName = "new_service_confirmed",
                        RetrieveAsIfPublished = true
                    };
                    var attResponse_Service = (RetrieveAttributeResponse)crmService.Execute(attReq_Service);
                    var attMetadata_Service = (EnumAttributeMetadata)attResponse_Service.AttributeMetadata;
                    var Service_Confirm_Label = attMetadata_Service.OptionSet.Options.Where(x => x.Value == serviceValue).FirstOrDefault().Label.UserLocalizedLabel.Label;

                    /*خواندن مقدار فیلد مجموع زمان خدمات*/
                    sum_Service = Convert.ToInt32(contract.new_sum_service_time);

                }

                /*بررسی شروط فروش*/
                if ((OptionSetValue)contract.new_sale_confirmed != null && saleValue == 100000000)
                {
                    if (contract.new_last_confirmed_date is null)
                    {
                        //خواندن تاریخ 
                        DateTime sale_LastDate = Convert.ToDateTime(contract.new_sale_confirmed_date);
                        DateTime sale_CreateOn = Convert.ToDateTime(contract.CreatedOn);
                        //int sale_Diff = Convert.ToInt32((sale_LastDate - sale_CreateOn).TotalHours);
                        int sale_Diff = Convert.ToInt32((sale_LastDate - sale_CreateOn).TotalMinutes);
                        int res = (sale_Diff + sum_Sale);

                        //contract.Attributes.Add(contract.new_last_confirmed_date.ToString(), sale_LastDate);
                        crmService.Update(new new_contract { Id = entity.Id, new_sum_sale_time = res, new_last_confirmed_date = sale_LastDate });
                    }
                    else
                    {
                        //خواندن تاریخ 
                        DateTime confirm_LastDate = Convert.ToDateTime(contract.new_last_confirmed_date);
                        DateTime sale_LastDate = Convert.ToDateTime(contract.new_sale_confirmed_date);
                        //int sale_Diff = Convert.ToInt32((sale_LastDate - confirm_LastDate).TotalHours);
                        int sale_Diff = Convert.ToInt32((sale_LastDate - confirm_LastDate).TotalMinutes);
                        int res = (sale_Diff + sum_Sale);

                        crmService.Update(new new_contract { Id = entity.Id, new_sum_sale_time = res, new_last_confirmed_date = sale_LastDate });
                    }
                }
                else if ((OptionSetValue)contract.new_sale_confirmed != null && (saleValue == 100000001 || saleValue == 100000002))
                {
                    if (contract.new_last_confirmed_date is null)
                    {
                        //خواندن تاریخ 
                        DateTime sale_LastDate = Convert.ToDateTime(contract.new_sale_confirmed_date);
                        DateTime sale_CreateOn = Convert.ToDateTime(contract.CreatedOn);
                        //int sale_Diff = Convert.ToInt32((sale_LastDate - sale_CreateOn).TotalHours);
                        int sale_Diff = Convert.ToInt32((sale_LastDate - sale_CreateOn).TotalMinutes);
                        int res = (sale_Diff + sum_Sale);

                        crmService.Update(new new_contract { Id = entity.Id, new_sum_sale_time = res, new_last_confirmed_date = sale_LastDate });
                        //return;
                    }
                    else
                    {
                        //خواندن تاریخ 
                        DateTime confirm_LastDate = Convert.ToDateTime(contract.new_last_confirmed_date);
                        DateTime sale_LastDate = Convert.ToDateTime(contract.new_sale_confirmed_date);
                        //int sale_Diff = Convert.ToInt32((confirm_LastDate - sale_LastDate).TotalHours);
                        int sale_Diff = Convert.ToInt32((sale_LastDate - confirm_LastDate).TotalMinutes);
                        int res = (sale_Diff + sum_Sale);

                        crmService.Update(new new_contract { Id = entity.Id, new_sum_sale_time = res, new_last_confirmed_date = sale_LastDate });
                        //return;
                    }
                }

                /*بررسی شروط مهندسی فروش*/
                if ((OptionSetValue)contract.new_sale_engineering_confirmed != null && sale_engineeringValue == 100000000)
                {
                    //خواندن تاریخ 
                    DateTime confirmed_LastDate = Convert.ToDateTime(contract.new_last_confirmed_date);
                    DateTime saleEngineering_LastDate = Convert.ToDateTime(contract.new_sale_engineering_confirmed_date);
                    //int saleEngineering_Diff = Convert.ToInt32((saleEngineering_LastDate - confirmed_LastDate).TotalHours);
                    int saleEngineering_Diff = Convert.ToInt32((saleEngineering_LastDate - confirmed_LastDate).TotalMinutes);
                    int res = (saleEngineering_Diff + sum_SaleEngineering);

                    crmService.Update(new new_contract { Id = entity.Id, new_sum_sale_engineering_time = res, new_last_confirmed_date = saleEngineering_LastDate });
                }
                else if ((OptionSetValue)contract.new_sale_engineering_confirmed != null && (sale_engineeringValue == 100000001 || sale_engineeringValue == 100000002))
                {
                    //خواندن تاریخ 
                    DateTime confirmed_LastDate = Convert.ToDateTime(contract.new_last_confirmed_date);
                    DateTime saleEngineering_LastDate = Convert.ToDateTime(contract.new_sale_engineering_confirmed_date);
                    //int saleEngineering_Diff = Convert.ToInt32((saleEngineering_LastDate - confirmed_LastDate).TotalHours);
                    int saleEngineering_Diff = Convert.ToInt32((saleEngineering_LastDate - confirmed_LastDate).TotalMinutes);
                    int res = (saleEngineering_Diff + sum_SaleEngineering);

                    crmService.Update(new new_contract { Id = entity.Id, new_sum_sale_engineering_time = res, new_last_confirmed_date = saleEngineering_LastDate });
                    //return;
                }

                /*بررسی شروط خدمات*/
                if ((OptionSetValue)contract.new_service_confirmed != null && serviceValue == 100000000)
                {
                    //خواندن تاریخ 
                    DateTime confirmed_LastDate = Convert.ToDateTime(contract.new_last_confirmed_date);
                    DateTime service_LastDate = Convert.ToDateTime(contract.new_service_confirmed_date);
                    //int service_Diff = Convert.ToInt32((service_LastDate - confirmed_LastDate).TotalHours);
                    int service_Diff = Convert.ToInt32((service_LastDate - confirmed_LastDate).TotalMinutes);
                    int res = (service_Diff + sum_Service);

                    crmService.Update(new new_contract { Id = entity.Id, new_sum_service_time = res, new_last_confirmed_date = service_LastDate });
                }
                else if ((OptionSetValue)contract.new_service_confirmed != null && (serviceValue == 100000001 || serviceValue == 100000002))
                {
                    //خواندن تاریخ 
                    DateTime confirmed_LastDate = Convert.ToDateTime(contract.new_last_confirmed_date);
                    DateTime service_LastDate = Convert.ToDateTime(contract.new_service_confirmed_date);
                    //int service_Diff = Convert.ToInt32((service_LastDate - confirmed_LastDate).TotalHours);
                    int service_Diff = Convert.ToInt32((service_LastDate - confirmed_LastDate).TotalMinutes);
                    int res = (service_Diff + sum_Service);

                    crmService.Update(new new_contract { Id = entity.Id, new_sum_service_time = res, new_last_confirmed_date = service_LastDate });
                    //return;
                }
            }
        }
    }
}
