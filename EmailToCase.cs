using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using WF.Xrm.Core;

namespace WF.SY.Plugins
{
    public class EmailToCase : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            //service variable
            CrmExecutionContext context = new CrmExecutionContext(serviceProvider, this.GetType(), false);

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                if (context.PrimaryEntityName != "email")
                { return; }

                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (entity.Contains("regardingobjectid"))
                    {
                        EntityReference regarding = (EntityReference)entity["regardingobjectid"];

                        if (regarding.LogicalName == "lead")
                        {
                            return; // don't run on leads
                        }
                    }

                    string subject = entity.Contains("subject") ? (string)entity["subject"] : string.Empty;
                    string pattern = @"^[\s]*([\w]+\s?:[\s]*)+"; // taken from CRM's "Smart Matching"
                    string replacement = string.Empty;
                    Regex regex = new Regex(pattern);

                    // ignore email replies and forwards (i.e. "RE:" and "FW:")
                    string subjectRegex = regex.Replace(subject, replacement);

                    List<Guid> partyIDs = new List<Guid>();
                    partyIDs.AddRange(GetPartyIDs(entity, "from", context));
                    partyIDs.AddRange(GetPartyIDs(entity, "to", context));
                    partyIDs.AddRange(GetPartyIDs(entity, "cc", context));
                    partyIDs.AddRange(GetPartyIDs(entity, "bcc", context));

                    // find open cases for any of the email parties, most recent first
                    QueryExpression q = new QueryExpression("incident");
                    q.ColumnSet = new ColumnSet("customerid", "title");
                    q.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); // active
                    q.AddOrder("createdon", OrderType.Descending);

                    FilterExpression childFilter = q.Criteria.AddFilter(LogicalOperator.Or);
                    foreach (Guid id in partyIDs)
                    {
                        childFilter.AddCondition("customerid", ConditionOperator.Equal, id);
                    }

                    EntityCollection cases = context.RetrieveMultiple(q);

                    if (cases.Entities.Count > 0)
                    {
                        bool attached = false; // email attached to correct case?

                        foreach (Entity c in cases.Entities)
                        {
                            if (!attached && c.Contains("title"))
                            {
                                string title = (string)c["title"];

                                if (title.Contains(subjectRegex))
                                {
                                    // attach the email to the appropriate case
                                    entity["regardingobjectid"] = new EntityReference()
                                    {
                                        LogicalName = c.LogicalName,
                                        Id = c.Id,
                                        Name = title
                                    };

                                    attached = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        entity["regardingobjectid"] = null;
                    }
                }
                catch (Exception ex)
                {
                    context.Trace("There was an error in the WF.Synopsys.Plugins.EmailToCase plugin: {0}", ex.ToString());
                    throw;
                }
            }
        }

        public List<Guid> GetPartyIDs(Entity email, string emailField, CrmExecutionContext context)
        {
            List<Guid> ids = new List<Guid>();

            if (email.Contains(emailField))
            {
                EntityCollection partyList = (EntityCollection)email[emailField];

                foreach (Entity p in partyList.Entities)
                {
                    if (p.Contains("partyid"))
                    {
                        EntityReference partyId = (EntityReference)p["partyid"];
                        ids.Add(partyId.Id);
                    }
                }
            }

            return ids;
        }

        public void RemoveFromAllQueues(Entity entity, CrmExecutionContext context)
        {
            QueryExpression q = new QueryExpression("queueitem");
            q.Criteria.AddCondition("objectid", ConditionOperator.Equal, entity.Id);

            var items = context.RetrieveMultiple(q);

            foreach (var i in items.Entities)
            {
                context.Delete(i.LogicalName, i.Id);
            }
        }
    }
}