using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//namespaces for d365 interaction
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using System.Globalization;
using System.Activities;
using Microsoft.Xrm.Sdk.Workflow;

namespace Customization_XRM.Base.WorkFlowBase
{
    public abstract class WorkflowBase : CodeActivity
    {
        internal string workflowClassName { get; set; }
        public WorkflowBase(Type workflowName)
        {
            this.workflowClassName = workflowName.ToString();
        }

        protected class LocalWorkflowExecution
        {
            internal IServiceProvider serviceProvider { get; set; }
            internal IOrganizationServiceFactory serviceFactory { get; set; }
            internal IOrganizationService orgService { get; set; }
            internal IWorkflowContext pluginContext { get; set; }
            internal ITracingService tracingService { get; set; }
            internal CodeActivityContext executeContext { get; set; }

            public LocalWorkflowExecution(CodeActivityContext serviceProvider)
            {
                if (serviceProvider == null)
                {
                    throw new InvalidPluginExecutionException("Invalied Service Provider");
                }
                this.pluginContext = (IWorkflowContext)serviceProvider.GetExtension<IWorkflowContext>();
                this.serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetExtension<IOrganizationServiceFactory>();
                this.orgService = this.serviceFactory.CreateOrganizationService(this.pluginContext.UserId);
                this.tracingService = (ITracingService)serviceProvider.GetExtension<ITracingService>();
                this.executeContext = serviceProvider;
            }

            public void TraceMessage(string message)
            {
                if (string.IsNullOrWhiteSpace(message))
                {
                    return;
                }
                if (this.pluginContext == null)
                {
                    this.tracingService.Trace("Invalied Plugin execution context");
                }
                else
                {
                    this.tracingService.Trace(string.Format(CultureInfo.InvariantCulture, "Correlated : {0}, UserId : {1}, message : {2}",
                                                                this.pluginContext.CorrelationId, this.pluginContext.UserId, message));
                }
            }
        }

        protected override void Execute(CodeActivityContext executeContext)
        {
            if (executeContext == null)
            {
                throw new InvalidPluginExecutionException("Service Provider is not initialized correctly.");
            }
            LocalWorkflowExecution localWorkFlowExecution = new LocalWorkflowExecution(executeContext);
            localWorkFlowExecution.TraceMessage(string.Format(CultureInfo.InvariantCulture, "Entered in {0}.Execute() method.", this.workflowClassName));
            try
            {
                ExecuteWorkFlowLogic(localWorkFlowExecution);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            finally
            {
                localWorkFlowExecution.TraceMessage(string.Format(CultureInfo.InvariantCulture, "Exit {0}.Execute method.", this.workflowClassName));
            }
        }

        //virtual method
        protected virtual void ExecuteWorkFlowLogic(LocalWorkflowExecution localWorkflowExecution)
        {
            //Plugin logic will be written in derived plugin class by overriding this method.
        }
    }
}
