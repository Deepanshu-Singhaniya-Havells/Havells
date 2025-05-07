using Microsoft.Extensions.Configuration;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using HavellsSync_ModelData.Common;
using ConfigurationManager = System.Configuration.ConfigurationManager;
using System.Configuration;

public class CrmService : ICrmService
{
    private readonly ServiceClient service;

    public CrmService(IConfiguration configuration)
    {

        Check.Argument.IsNotNull(nameof(configuration), configuration);

        var clientId = configuration["CRMSettings:ClientId"];
        var clientSecret = configuration["CRMSettings:ClientSecret"];
        var orgUrl = configuration["CRMSettings:OrgUrl"];
        var conn = ConfigurationManager.ConnectionStrings["key"];
        service = new ServiceClient($@"AuthType=ClientSecret;url={orgUrl};ClientId={clientId};ClientSecret={clientSecret}");
    }

    /// <summary>
    /// The method retrieves record for expression query.
    /// </summary>
    /// <param name="query"></param>
    /// <returns></returns>
    public EntityCollection RetrieveMultiple(FetchExpression query)
    {
        return service.RetrieveMultiple(query);
    }

    /// <summary>
    /// The method retrieves record for given entityname 
    /// where queryattribute is equal to query attributes values.
    /// </summary>
    /// <param name="entityName"></param>
    /// <param name="queryAttributes"></param>
    /// <param name="queryAttributeValues"></param>
    /// <param name="columns"></param>
    /// <returns></returns>
    public EntityCollection RetrieveMultiple(string entityName, string[] queryAttributes, object[] queryAttributeValues, ColumnSet columns)
    {
        QueryExpression query = new QueryExpression(entityName);
        if (queryAttributes != null && queryAttributes.Length > 0 && queryAttributeValues != null && queryAttributeValues.Length > 0)
        {
            FilterExpression filter = new FilterExpression();
            for (int i = 0; i < queryAttributes.Length; i++)
            {
                ConditionExpression condition = new ConditionExpression(queryAttributes[i], ConditionOperator.Equal, queryAttributeValues[i]);
                filter.AddCondition(condition);
            }
            query.Criteria = filter;
        }
        if (columns != null)
            query.ColumnSet = columns;

        return service.RetrieveMultiple(query);
    }


    /// <summary>
    /// The method retrieve the result set for given entity
    /// based on conditions provided.
    /// </summary>
    /// <param name="entityName"></param>
    /// <param name="conditionExpressions"></param>
    /// <param name="columns"></param>
    /// <returns></returns>
    public EntityCollection RetrieveMultiple(string entityName, List<ConditionExpression> conditionExpressions, ColumnSet columns)
    {
        QueryExpression query = new QueryExpression(entityName);

        foreach (var conditionExpression in conditionExpressions)
        {
            query.Criteria.AddCondition(conditionExpression);
        }


        if (columns != null)
            query.ColumnSet = columns;

        return service.RetrieveMultiple(query);
    }


    /// <summary>
    /// The method retrieves result with link entity
    /// </summary>
    /// <param name="fromEntityName"></param>
    /// <param name="toEntityName"></param>
    /// <param name="linkFromAttribute"></param>
    /// <param name="linkToAttribute"></param>
    /// <param name="alias"></param>
    /// <param name="fromConditionExpressions"></param>
    /// <param name="toConditionExpressions"></param>
    /// <param name="fromColumnSet"></param>
    /// <param name="toColumnSet"></param>
    /// <param name="joinOperator"></param>
    /// <returns></returns>
    public EntityCollection RetrieveMultiple(string fromEntityName, string toEntityName, string linkFromAttribute, string linkToAttribute, string alias,
        List<ConditionExpression> fromConditionExpressions, List<ConditionExpression> toConditionExpressions,
        ColumnSet fromColumnSet, ColumnSet toColumnSet, JoinOperator joinOperator)
    {
        QueryExpression query = new QueryExpression(fromEntityName);
        query.ColumnSet = fromColumnSet;

        if (fromConditionExpressions != null)
        {
            foreach (var fromConditionExpression in fromConditionExpressions)
            {
                query.Criteria.AddCondition(fromConditionExpression);
            }
        }

        LinkEntity toEntity = new LinkEntity(fromEntityName, toEntityName, linkFromAttribute, linkToAttribute, joinOperator);
        toEntity.EntityAlias = alias;
        toEntity.Columns = toColumnSet;

        if (toConditionExpressions != null)
        {
            foreach (var conditionExpression in toConditionExpressions)
            {
                toEntity.LinkCriteria.AddCondition(conditionExpression);
            }
        }
        query.LinkEntities.Add(toEntity);
        return service.RetrieveMultiple(query);
    }

    /// <summary>
    /// The method retrieves the result for given query
    /// </summary>
    /// <param name="query"></param>
    /// <returns></returns>
    public EntityCollection RetrieveMultiple(QueryExpression query)
    {
        return service.RetrieveMultiple(query);
    }

    /// <summary>
    /// The method creates a entity record in crm
    /// </summary>
    /// <param name="e"></param>
    /// <returns></returns>
    public Guid Create(Entity e)
    {
        return service.Create(e);
    }

    /// <summary>
    /// The method updates the record in crm
    /// </summary>
    /// <param name="e"></param>
    public void Update(Entity e)
    {
        service.Update(e);
    }

    public UpsertResponse Upsert(Entity e)
    {
        var request = new UpsertRequest()
        {
            Target = e
        };

        var response = (UpsertResponse)service.Execute(request);

        return response;
    }

    /// <summary>
    /// The method retrieve the record for given entity for given entity id
    /// </summary>
    /// <param name="entityname"></param>
    /// <param name="entityid"></param>
    /// <param name="column"></param>
    /// <returns></returns>
    public Entity Retrieve(string entityname, Guid entityid, ColumnSet column)
    {
        return service.Retrieve(entityname, entityid, column);
    }

    /// <summary>
    /// The method executes any crm request and 
    /// send organizationresponse back to calling method.
    /// </summary>
    /// <param name="request"></param>
    /// <returns></returns>
    public OrganizationResponse Execute(OrganizationRequest request)
    {
        return service.Execute(request);
    }

    /// <summary>
    /// The method delete a record from given entity
    /// </summary>
    /// <param name="entityname">The entity name.</param>
    /// <param name="entityid">The entity id.</param>
    public void Delete(string entityname, Guid entityid)
    {
        service.Delete(entityname, entityid);
    }

    public EntityReference GetEntityReference(string entityName, string primaryFieldName, string primaryFieldValue)
    {
        EntityReference entityReference = null;
        try
        {
            QueryExpression queryExp = new QueryExpression(entityName);
            queryExp.ColumnSet = new ColumnSet(false);
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition(primaryFieldName, ConditionOperator.Equal, primaryFieldValue);
            EntityCollection entityCollection = service.RetrieveMultiple(queryExp);
            if (entityCollection.Entities.Count > 0)
            {
                entityReference = entityCollection.Entities[0].ToEntityReference();
            }
        }
        catch { }
        return entityReference;
    }

    /// <summary>
    /// To get option set values.
    /// </summary>
    /// <param name="entityName">The entity name.</param>
    /// <param name="attributeName">The attribute name.</param>
    /// <returns></returns>
    public KeyValuePair<int, string>[] GetOptionSetValues(string entityName, string attributeName)
    {
        var attributeRequest = new RetrieveAttributeRequest
        {
            EntityLogicalName = entityName,
            LogicalName = attributeName,
            RetrieveAsIfPublished = true
        };

        var attributeResponse = (RetrieveAttributeResponse)service.Execute(attributeRequest);
        var attributeMetadata = (EnumAttributeMetadata)attributeResponse.AttributeMetadata;

        return attributeMetadata.OptionSet.Options.
            Where(x => x.Value.HasValue).
            Select(x => new KeyValuePair<int, string>(x.Value.Value, x.Label.UserLocalizedLabel.Label)).ToArray();
    }
}
