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

namespace Customization_XRM.Base.PluginBase
{
    public abstract class PluginBase : IPlugin
    {
        internal string pluginClassName { get; set; }

        public PluginBase(Type pluginName)
        {
            this.pluginClassName = pluginName.ToString();
        }
        protected class LocalPluginExecution
        {
            internal IServiceProvider serviceProvider { get; set; }
            internal IOrganizationServiceFactory serviceFactory { get; set; }
            internal IOrganizationService orgService { get; set; }
            internal IPluginExecutionContext pluginContext { get; set; }
            internal ITracingService tracingService { get; set; }

            public LocalPluginExecution(IServiceProvider serviceProvider)
            {
                if (serviceProvider == null)
                {
                    throw new InvalidPluginExecutionException("Invalied Service Provider");
                }
                this.pluginContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                this.serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                this.orgService = this.serviceFactory.CreateOrganizationService(this.pluginContext.UserId);
                this.tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
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
        public void Execute(IServiceProvider serviceProvider)
        {
            if (serviceProvider == null)
            {
                throw new InvalidPluginExecutionException("Service Provider is not initialized correctly.");
            }
            LocalPluginExecution localPluginExecution = new LocalPluginExecution(serviceProvider);
            localPluginExecution.TraceMessage(string.Format(CultureInfo.InvariantCulture, "Entered in {0}.Execute() method.", this.pluginClassName));
            try
            {
                ExecutePluginLogic(localPluginExecution);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            finally
            {
                localPluginExecution.TraceMessage(string.Format(CultureInfo.InvariantCulture, "Exit {0}.Execute method.", this.pluginClassName));
            }
        }

        //virtual method
        protected virtual void ExecutePluginLogic(LocalPluginExecution localPluginExecution)
        {
            //Plugin logic will be written in derived plugin class by overriding this method.
        }
    }
}
