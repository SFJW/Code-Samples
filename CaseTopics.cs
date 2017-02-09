using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using WF.Xrm.Core;

namespace AV.MC.Plugins
{
    public class CaseTopics : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            //service variable
            CrmExecutionContext context = new CrmExecutionContext(serviceProvider, this.GetType(), false);
            EntityReference targetEntity = null;
            string relationshipName = string.Empty;
            EntityReferenceCollection relatedEntities = null;
            EntityReference relatedEntity = null;

            if (context.MessageName == "Associate" || context.MessageName == "Disassociate")
            {
                // Get the "Relationship" Key from context
                if (context.InputParameters.Contains("Relationship"))
                {
                    relationshipName = context.InputParameters["Relationship"].ToString();
                }

                context.Trace("Relationship: {0}", relationshipName);
                
                // Check the "Relationship Name" with your intended one; I don't know why the relationship has a period, but it does
                if (relationshipName != "mc_N_to_N_topic_incident.")
                { return; } 

                // Get Entity 1 reference from "Target" Key from context
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                {
                    targetEntity = (EntityReference)context.InputParameters["Target"];
                }

                // Get Entity 2 reference from "RelatedEntities" Key from context
                if (context.InputParameters.Contains("RelatedEntities") && context.InputParameters["RelatedEntities"] is EntityReferenceCollection)
                {
                    relatedEntities = context.InputParameters["RelatedEntities"] as EntityReferenceCollection;
                    relatedEntity = relatedEntities[0];
                }
                
                if (relatedEntity.LogicalName != "incident" && relatedEntity.LogicalName != "mc_topic")
                { return; }

                try
                {
                    Guid caseId = relatedEntity.LogicalName == "incident" ? caseId = relatedEntity.Id : caseId = targetEntity.Id;

                    string entity1 = "mc_topic";
                    string entity2 = "incident";
                    string relationshipEntityName = "mc_n_to_n_topic_incident";

                    QueryExpression query = new QueryExpression(entity1);
                    query.ColumnSet = new ColumnSet("mc_topicname");
                    query.AddOrder("mc_topicname", OrderType.Ascending);

                    LinkEntity linkEntity1 = new LinkEntity(entity1, relationshipEntityName, "mc_topicid", "mc_topicid", JoinOperator.Inner);
                    LinkEntity linkEntity2 = new LinkEntity(relationshipEntityName, entity2, "incidentid", "incidentid", JoinOperator.Inner);
                    linkEntity1.LinkEntities.Add(linkEntity2);
                    query.LinkEntities.Add(linkEntity1);

                    // filter by the current case - show topics for case
                    linkEntity2.LinkCriteria = new FilterExpression();
                    linkEntity2.LinkCriteria.AddCondition(new ConditionExpression("incidentid", ConditionOperator.Equal, caseId));
                    
                    // build the list of topics
                    EntityCollection collRecords = context.RetrieveMultiple(query);
                    List<string> topics = new List<string>();
                    
                    foreach (var t in collRecords.Entities)
                    {
                        if (t.Contains("mc_topicname"))
                        {
                            string topic = (string)t["mc_topicname"];
                            topics.Add(topic);
                        }
                    }

                    // concatenate the list of topics
                    string topicsConcat = String.Join("\n", topics.ToArray());
                    
                    // update the case
                    Entity relatedCase = new Entity("incident");
                    relatedCase.Id = caseId;
                    relatedCase["mc_topics"] = topicsConcat;
                    context.Update(relatedCase);
                }
                catch (Exception ex)
                {
                    context.Trace("There was an error in the Avtex.MC.Plugins.CaseTopics plugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}
