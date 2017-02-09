using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using WF.Xrm.Core;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;

namespace WF.FBG.Plugins
{
    public class OptionSet
    {
        string entityName { get; set; }
        string optionSet { get; set; }
        CrmExecutionContext context { get; set; }

        public OptionSet(string entityLogicalName, string optionSetAttribute, CrmExecutionContext executionContext)
        {
            entityName = entityLogicalName;
            optionSet = optionSetAttribute;
            context = executionContext;
        }

        public OptionMetadata[] GetOptions()
        {
            RetrieveAttributeRequest request = new RetrieveAttributeRequest();
            request.EntityLogicalName = entityName;
            request.LogicalName = optionSet;
            request.RetrieveAsIfPublished = true;

            // execute the request
            RetrieveAttributeResponse response = (RetrieveAttributeResponse)context.Execute(request);

            // get the options
            PicklistAttributeMetadata picklistMetadata = (PicklistAttributeMetadata)response.AttributeMetadata;
            OptionMetadata[] options = picklistMetadata.OptionSet.Options.ToArray();

            return options;
        }
        public string GetLabel(int value)
        {
            RetrieveAttributeRequest request = new RetrieveAttributeRequest();
            request.EntityLogicalName = entityName;
            request.LogicalName = optionSet;
            request.RetrieveAsIfPublished = true;

            // execute the request
            RetrieveAttributeResponse response = (RetrieveAttributeResponse)context.Execute(request);

            // get the options
            PicklistAttributeMetadata picklistMetadata = (PicklistAttributeMetadata)response.AttributeMetadata;
            OptionMetadata[] options = picklistMetadata.OptionSet.Options.ToArray();

            string label = String.Empty;

            foreach (OptionMetadata option in options)
            {
                if (option.Value == value)
                {
                    label = option.Label.UserLocalizedLabel.Label;
                    return label;
                }
            }

            // didn't find a match
            return label;
        }
        public int GetValue(string label)
        {
            RetrieveAttributeRequest request = new RetrieveAttributeRequest();
            request.EntityLogicalName = entityName;
            request.LogicalName = optionSet;
            request.RetrieveAsIfPublished = true;

            // execute the request
            RetrieveAttributeResponse response = (RetrieveAttributeResponse)context.Execute(request);

            // get the options
            PicklistAttributeMetadata picklistMetadata = (PicklistAttributeMetadata)response.AttributeMetadata;
            OptionMetadata[] options = picklistMetadata.OptionSet.Options.ToArray();

            int value = 0;

            foreach (OptionMetadata option in options)
            {
                if (option.Label.UserLocalizedLabel.Label.ToLower() == label.ToLower())
                {
                    value = option.Value.Value;
                    return value;
                }
            }

            // didn't find a match
            return value;
        }
    }

    public static class ZipCodeMapping
    {
        public static void Launch(Entity entity, Entity postImage, int source, CrmExecutionContext context)
        {
            // get the zip code
            string zipCode = String.Empty;

            if (entity.Contains("address1_postalcode"))
            {
                zipCode = (string)entity["address1_postalcode"];
            }
            else if (postImage.Contains("address1_postalcode"))
            {
                zipCode = (string)postImage["address1_postalcode"];
            }

            // get the state
            OptionSetValue state = null;

            if (entity.Contains("fbg_state"))
            {
                state = (OptionSetValue)entity["fbg_state"];
            }
            else if (postImage.Contains("fbg_state"))
            {
                state = (OptionSetValue)postImage["fbg_state"];
            }

            context.Trace("Zip Code: {0}\nState: {1}", zipCode, state != null ? state.Value.ToString() : string.Empty);
            
            // has a zip code and state
            if (!String.IsNullOrEmpty(zipCode) && state != null)
            {
                context.Trace("Has a zip code and state.");

                bool foundZip = MatchZip(entity, zipCode, source, context);

                if (!foundZip)
                {
                    bool foundState = MatchState(entity, state, source, context);

                    if (!foundState)
                    {
                        // default new owner is the default user
                        Owner.SetDefault(entity, context);
                    }
                }
            }
            // has a zip code, no state
            else if (!String.IsNullOrEmpty(zipCode) && state == null)
            {
                context.Trace("Has a zip code but no state.");

                bool foundZip = MatchZip(entity, zipCode, source, context);

                if (!foundZip)
                {
                    // default new owner is the default user
                    Owner.SetDefault(entity, context);
                }
            }
            // has a state, no zip code
            else if (String.IsNullOrEmpty(zipCode) && state != null)
            {
                context.Trace("Has a state but no zip code.");

                bool foundState = MatchState(entity, state, source, context);

                if (!foundState)
                {
                    // default new owner is the default user
                    Owner.SetDefault(entity, context);
                }
            }
            // does not have a zip code or state
            else
            {
                context.Trace("Does not have a zip code or state.");

                // default new owner is the default user
                Owner.SetDefault(entity, context);
            }
        }
        public static EntityCollection GetZipCodeMappings(ConditionExpression condition, int source, CrmExecutionContext context)
        {
            QueryExpression query = new QueryExpression("wf_zipcodemapping");
            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddCondition(condition);
            query.Criteria.AddCondition("wf_source", ConditionOperator.Equal, source);

            EntityCollection mappings = context.RetrieveMultiple(query);

            return mappings;
        }
        public static bool EvaluateMappings(Entity entity, EntityCollection mappings, CrmExecutionContext context)
        {
            bool matchFound = false;

            if (mappings.Entities.Count > 0)
            {
                matchFound = true;

                // get the first result
                Entity mapping = mappings.Entities.First<Entity>();

                // the new owner is the user in charge of the zip code
                EntityReference zipCodeUser = (EntityReference)mapping["wf_user"];

                Owner.SetNewOwner(entity, zipCodeUser, context);
            }

            return matchFound;
        }
        public static bool MatchZip(Entity entity, string zipCode, int source, CrmExecutionContext context)
        {
            // remove all non-digit characters
            string allDigits = new string(zipCode.Where(c => char.IsDigit(c)).ToArray());

            // pad the end of the zip code with zeros if the zip is less than five digits; select the first five digits
            string firstFiveDigits = allDigits.PadRight(5, '0').Substring(0, 5);

            ConditionExpression zipCondition = new ConditionExpression("wf_name", ConditionOperator.In, 
                firstFiveDigits,
                firstFiveDigits.Substring(0, 4),
                firstFiveDigits.Substring(0, 3),
                firstFiveDigits.Substring(0, 2),
                firstFiveDigits.Substring(0, 1),
                firstFiveDigits.Substring(0, 4) + "+",
                firstFiveDigits.Substring(0, 3) + "+",
                firstFiveDigits.Substring(0, 2) + "+",
                firstFiveDigits.Substring(0, 1) + "+");

            // get the records
            EntityCollection mappings = GetZipCodeMappings(zipCondition, source, context);

            // keep going if at least one record is returned
            bool matchFound = EvaluateMappings(entity, mappings, context);

            return matchFound;
        }
        public static bool MatchState(Entity entity, OptionSetValue state, int source, CrmExecutionContext context)
        {
            // find a match between the source record's state and zip code mapping's state
            ConditionExpression stateCondition = new ConditionExpression("wf_state", ConditionOperator.Equal, state.Value);

            // get the records
            EntityCollection mappings = GetZipCodeMappings(stateCondition, source, context);

            // keep going if at least one record is returned
            bool matchFound = EvaluateMappings(entity, mappings, context);

            return matchFound;
        }
    }

    public static class Owner
    {
        public static void SetDefault(Entity entity, CrmExecutionContext context)
        {
            Guid userId = getUserId("CRM Admin", context);

            if (!userId.Equals(Guid.Empty))
            {
                EntityReference defaultUser = new EntityReference("systemuser", userId);

                // default new owner is the default user
                SetNewOwner(entity, defaultUser, context);
            }
        }
        public static void SetNewOwner(Entity entity, EntityReference newOwner, CrmExecutionContext context)
        {
            context.Trace("New owner: {0}", newOwner.Name);

            // set the owner
            AssignRequest assignRequest = new AssignRequest();
            assignRequest.Assignee = newOwner;
            assignRequest.Target = new EntityReference(entity.LogicalName, entity.Id);
            context.Execute(assignRequest);
        }
        public static Guid getTeamId(string teamName, CrmExecutionContext context)
        {
            QueryExpression q = new QueryExpression("team");
            q.ColumnSet = new ColumnSet(false);
            q.Criteria.AddCondition("name", ConditionOperator.Equal, teamName);

            var teams = context.RetrieveMultiple(q);

            if (teams.Entities.Count > 0)
            {
                Entity firstTeam = teams.Entities.First();

                return firstTeam.Id;
            }

            return Guid.Empty;
        }
        public static Guid getUserId(string userFullName, CrmExecutionContext context)
        {
            QueryExpression q = new QueryExpression("systemuser");
            q.ColumnSet = new ColumnSet(false);
            q.Criteria.AddCondition("fullname", ConditionOperator.Equal, userFullName);

            var users = context.RetrieveMultiple(q);

            if (users.Entities.Count > 0)
            {
                Entity firstUser = users.Entities.First();

                return firstUser.Id;
            }

            return Guid.Empty;
        }
    }

    public class AssignLead : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            //service variable
            CrmExecutionContext context = new CrmExecutionContext(serviceProvider, this.GetType(), false);

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                try
                {
                    Entity postImage = (Entity)context.PostEntityImages["Target"];
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (entity.LogicalName != "lead")
                    { return; }

                    // only run on open leads
                    if (postImage.Contains("statecode") && ((OptionSetValue)postImage["statecode"]).Value != 0)
                    { return; }

                    // ONVIA Lead
                    if (postImage.Contains("wf_source"))
                    {
                        string sourceLowercase = ((string)postImage["wf_source"]).ToLower();
                        int sourceVal = 2; // SCA = 1, Davis Bacon = 2

                        if (sourceLowercase.Contains("sca"))
                        {
                            sourceVal = 1;
                        }
                        else if (sourceLowercase.Contains("davis bacon") || sourceLowercase.Contains("davis") || sourceLowercase.Contains("bacon"))
                        {
                            sourceVal = 2;
                        }
                        else
                        {
                            return;
                        }

                        ZipCodeMapping.Launch(entity, postImage, sourceVal, context);
                    }
                }
                catch (Exception ex)
                {
                    context.Trace("There was an error in the AssignLead plugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }

    /*public class AssignAccount : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            //service variable
            CrmExecutionContext context = new CrmExecutionContext(serviceProvider, this.GetType(), false);

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                try
                {
                    Entity postImage = (Entity)context.PostEntityImages["Target"];
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (entity.LogicalName != "account")
                    { return; }

                    ZipCodeMapping.Launch(entity, postImage, context);
                }
                catch (Exception ex)
                {
                    context.Trace("There was an error in the AssignLead plugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }*/
}
