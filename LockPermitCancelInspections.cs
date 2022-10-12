using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Sdk;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk.Query;

namespace ContosoPackageproject
{

    public class LockPermitCancelInspections : PluginBase
    {
        public LockPermitCancelInspections(string unsecureConfiguration, string secureConfiguration): base(typeof(PreOperationPermitCreate))
        {

        }
        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null) {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            var permitEntityRef = localPluginContext.PluginExecutionContext.InputParameters["Target"] as EntityReference;
            Entity permitEntity = new Entity(permitEntityRef.LogicalName, permitEntityRef.Id);
            localPluginContext.Trace("Updating Permit Id : " + permitEntityRef.Id);
            permitEntity["statuscode"] = new OptionSetValue(330650000);
            localPluginContext.PluginUserService.Update(permitEntity);
            localPluginContext.Trace("Updated Permit Id " + permitEntityRef.Id);

            QueryExpression qe = new QueryExpression();
            qe.EntityName = "contoso_inspection";
            qe.ColumnSet = new ColumnSet("statuscode");
            ConditionExpression condition = new ConditionExpression();
            condition.Operator = ConditionOperator.Equal;
            condition.AttributeName = "contoso_permit";
            condition.Values.Add(permitEntityRef.Id);
            qe.Criteria = new FilterExpression(LogicalOperator.And);
            qe.Criteria.Conditions.Add(condition);

            localPluginContext.Trace("Retrieving inspections for Permit Id " + permitEntityRef.Id);
            var inspectionsResult = localPluginContext.PluginUserService.RetrieveMultiple(qe);
            localPluginContext.Trace("Retrievied " + inspectionsResult.TotalRecordCount + " inspection records");

            int canceledInspectionsCount = 0;
            foreach (var inspection in inspectionsResult.Entities)
            {
                var currentValue = inspection.GetAttributeValue<OptionSetValue>("statuscode");
                if (currentValue.Value == 1 || currentValue.Value == 330650000)
                {
                    canceledInspectionsCount++;
                    inspection["statuscode"] = new OptionSetValue(330650003);
                    localPluginContext.Trace("Canceling inspection Id : " + inspection.Id);
                    localPluginContext.PluginUserService.Update(inspection);
                    localPluginContext.Trace("Canceled inspection Id : " + inspection.Id);
                }
            }

            if (canceledInspectionsCount > 0)
            {
                localPluginContext.PluginExecutionContext.OutputParameters["CanceledInspectionsCount"] = canceledInspectionsCount + " Inspections were canceled";
            }

            if (localPluginContext.PluginExecutionContext.InputParameters.ContainsKey("Reason"))
            {
                localPluginContext.Trace("building a note reocord");
                Entity note = new Entity("annotation");
                note["subject"] = "Permit Locked";
                note["notetext"] = "Reason for locking this permit: " + localPluginContext.PluginExecutionContext.InputParameters["Reason"];
                note["objectid"] = permitEntityRef;
                note["objecttypecode"] = permitEntityRef.LogicalName;

                localPluginContext.Trace("Creating a note reocord");
                var createdNoteId = localPluginContext.PluginUserService.Create(note);
                if (createdNoteId != Guid.Empty) localPluginContext.Trace("Note record was created");
            }

        }
    }
}
