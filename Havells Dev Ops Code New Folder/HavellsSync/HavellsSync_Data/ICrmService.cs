using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using HavellsSync_ModelData;

public interface ICrmService
{
    /// <summary>
    /// The method retrieves record for expression query.
    /// </summary>
    /// <param name="query"></param>
    /// <returns></returns>
    EntityCollection RetrieveMultiple(FetchExpression query);

    /// <summary>
    /// The method retrieves record for given entityname 
    /// where queryattribute is equal to query attributes values.
    /// </summary>
    /// <param name="entityName"></param>
    /// <param name="queryAttributes"></param>
    /// <param name="queryAttributeValues"></param>
    /// <param name="columns"></param>
    /// <returns></returns>
    EntityCollection RetrieveMultiple(string entityName, string[] queryAttributes, object[] queryAttributeValues, ColumnSet columns);


    /// <summary>
    /// The method retrieve the result set for given entity
    /// based on conditions provided.
    /// </summary>
    /// <param name="entityName"></param>
    /// <param name="conditionExpressions"></param>
    /// <param name="columns"></param>
    /// <returns></returns>
    EntityCollection RetrieveMultiple(string entityName, List<ConditionExpression> conditionExpressions, ColumnSet columns);


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
    EntityCollection RetrieveMultiple(string fromEntityName, string toEntityName, string linkFromAttribute, string linkToAttribute, string alias,
    List<ConditionExpression> fromConditionExpressions, List<ConditionExpression> toConditionExpressions,
        ColumnSet fromColumnSet, ColumnSet toColumnSet, JoinOperator joinOperator);

    /// <summary>
    /// The method retrieves the result for given query
    /// </summary>
    /// <param name="query"></param>
    /// <returns></returns>
    EntityCollection RetrieveMultiple(QueryExpression query);

    /// <summary>
    /// The method creates a entity record in crm
    /// </summary>
    /// <param name="e"></param>
    /// <returns></returns>
    Guid Create(Entity e);

    /// <summary>
    /// The method updates the record in crm
    /// </summary>
    /// <param name="e"></param>
    void Update(Entity e);

    /// <summary>
    /// To delete a record from a given entity.
    /// </summary>
    /// <param name="entityname">The entity name.</param>
    /// <param name="entityid">The entity id.</param>
    void Delete(string entityname, Guid entityid);

    UpsertResponse Upsert(Entity e);

    /// <summary>
    /// The method retrieve the record for given entity for given entity id
    /// </summary>
    /// <param name="entityname"></param>
    /// <param name="entityid"></param>
    /// <param name="column"></param>
    /// <returns></returns>
    Entity Retrieve(string entityname, Guid entityid, ColumnSet column);

    /// <summary>
    /// The method executes any crm request and 
    /// send organizationresponse back to calling method.
    /// </summary>
    /// <param name="request"></param>
    /// <returns></returns>
    OrganizationResponse Execute(OrganizationRequest request);

    /// <summary>
    /// To get option set values.
    /// </summary>
    /// <param name="entityName">The entity name.</param>
    /// <param name="attributeName">The attribute name.</param>
    /// <returns></returns>
    KeyValuePair<int, string>[] GetOptionSetValues(string entityName, string attributeName);

    EntityReference GetEntityReference(string entityName, string primaryFieldName, string primaryFieldValue);
}
