// <copyright file="PostUpdateProjectTaskCommon.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>
//// <author></author>
//// <date>4/24/2020 1:30:47 PM</date>
//// <summary>Implements the Post updation of Project Task</summary>

namespace MCS.PSA.Plugins
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using MCS.PSA.Plugins.BusinessLogic;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;

    /// <summary>
    /// PostUpdateProjectTaskCommon class.
    /// </summary>
    public class PostUpdateProjectTaskCommon : IPlugin
    {
        /// <summary>
        /// Enabled For Snapshot.
        /// </summary>
        /// <param name="targetEntity">target Entity.</param>
        /// <param name="service">Organization service.</param>
        /// <param name="tracingService">tracing service.</param>
        /// <param name="preImage">pre Image.</param>
        /// <param name="postImage">post Image.</param>
        /// <param name="presentStartDate">present StartDate.</param>
        /// <param name="presentEndDate">present EndDate.</param>
        /// <param name="previousStartDate">previous StartDate.</param>
        /// <param name="previousEndDate">previous EndDate.</param>
        /// <param name="project">entity reference project.</param>
        public static void EnabledForSnapshot(Entity targetEntity, IOrganizationService service, ITracingService tracingService, Entity preImage, Entity postImage, DateTime presentStartDate, DateTime presentEndDate, DateTime previousStartDate, DateTime previousEndDate, EntityReference project)
        {
            if (targetEntity.Contains(Common.Model.ProjectTask.Milestone) == false && preImage.Contains(Common.Model.ProjectTask.Milestone) == false)
            {
                tracingService.Trace("Milesone Not Present");
                return;
            }

            if (targetEntity.Contains(Common.Model.ProjectTask.ScheduledStart) && targetEntity.Attributes[Common.Model.ProjectTask.ScheduledStart] != null && postImage.Contains(Common.Model.ProjectTask.Milestone) == true)
            {
                Guid milestoneId = ((EntityReference)postImage.Attributes[Common.Model.ProjectTask.Milestone]).Id;
                Entity milestoneDetails = service.Retrieve(Common.Model.MilestoneCategory.LogicalName, milestoneId, new ColumnSet(Common.Model.MilestoneCategory.Criteria));
                if (milestoneDetails.Contains(Common.Model.MilestoneCategory.Criteria))
                {
                    string milestoneCriteria = ((OptionSetValue)milestoneDetails.Attributes[Common.Model.MilestoneCategory.Criteria]).Value.ToString(CultureInfo.CurrentCulture);
                    if (milestoneCriteria == Common.Model.MilestoneCategory.MilestoneCriteriaStart)
                    {
                        presentStartDate = (DateTime)targetEntity.Attributes[Common.Model.ProjectTask.ScheduledStart];
                        tracingService.Trace(FormattableString.Invariant($"PresentStart {presentStartDate.ToString(CultureInfo.InvariantCulture)}"));
                        BusinessLogic.ProjectTask.CompareValuesAndSetFlag(service, presentStartDate, tracingService, targetEntity.Id, project.Id);
                    }
                }
            }

            if (targetEntity.Contains(Common.Model.ProjectTask.ScheduledEnd) && targetEntity.Attributes[Common.Model.ProjectTask.ScheduledEnd] != null)
            {
                Guid milestoneId = ((EntityReference)postImage.Attributes[Common.Model.ProjectTask.Milestone]).Id;
                Entity milestoneDetails = service.Retrieve(Common.Model.MilestoneCategory.LogicalName, milestoneId, new ColumnSet(Common.Model.MilestoneCategory.Criteria));
                if (milestoneDetails.Contains(Common.Model.MilestoneCategory.Criteria))
                {
                    string milestoneCriteria = ((OptionSetValue)milestoneDetails.Attributes[Common.Model.MilestoneCategory.Criteria]).Value.ToString(CultureInfo.CurrentCulture);
                    if (milestoneCriteria == Common.Model.MilestoneCategory.MilestoneCriteriaEnd)
                    {
                        presentEndDate = (DateTime)targetEntity.Attributes[Common.Model.ProjectTask.ScheduledEnd];
                        tracingService.Trace(FormattableString.Invariant($"PresentEnd  {presentEndDate.ToString(CultureInfo.InvariantCulture)}"));
                        BusinessLogic.ProjectTask.CompareValuesAndSetFlag(service, presentEndDate, tracingService, targetEntity.Id, project.Id);
                    }
                }
            }
        }

        /// <summary>
        /// Updates the Actual Start Date and End Date of Project Tasks to Parent Tasks.
        /// </summary>
        /// <param name="postImage">Post Image.</param>
        /// <param name="tracingService">Tracing Service.</param>
        /// <param name="targetEntity">Target Entity.</param>
        /// <param name="service">For Service.</param>
        public static void UpdateActualStartandActualEndDate(Entity postImage, ITracingService tracingService, Entity targetEntity, IOrganizationService service)
        {
            tracingService.Trace("UpdateActualStartandActualEndDate method is being executed.");
            DateTime? parentStartDate = null;
            DateTime? parentEndDate = null;

            Guid parentTaskID = ((EntityReference)postImage.Attributes[Common.Model.ProjectTask.ParentTask]).Id;
            tracingService.Trace(FormattableString.Invariant($"Parent Task ID: ") + parentTaskID);
            Entity projectTaskEntity = new Entity(Common.Model.ProjectTask.LogicalName, parentTaskID);

            QueryExpression listofChildTaskswithStartDate = new QueryExpression
            {
                EntityName = Common.Model.ProjectTask.LogicalName,
                ColumnSet = new ColumnSet(Common.Model.ProjectTask.ActualStart, Common.Model.ProjectTask.ActualEnd),
            };
            listofChildTaskswithStartDate.Criteria.AddCondition(Common.Model.ProjectTask.ParentTask, ConditionOperator.Equal, parentTaskID);
            listofChildTaskswithStartDate.Criteria.AddCondition(Common.Model.ProjectTask.ActualStart, ConditionOperator.NotNull);
            listofChildTaskswithStartDate.NoLock = true;
            EntityCollection startDateChildTaskCollection = service.RetrieveMultiple(listofChildTaskswithStartDate);
            tracingService.Trace(FormattableString.Invariant($"Start Date Child Task Count: {startDateChildTaskCollection.Entities.Count}"));
            if (startDateChildTaskCollection.Entities.Any())
            {
                parentStartDate = startDateChildTaskCollection.Entities.OrderBy(stdate => stdate.Attributes[Common.Model.ProjectTask.ActualStart]).FirstOrDefault().GetAttributeValue<DateTime>(Common.Model.ProjectTask.ActualStart);
                tracingService.Trace("Start Date: " + parentStartDate);
            }

            QueryExpression listofIncompleteChildTasks = new QueryExpression
            {
                EntityName = Common.Model.ProjectTask.LogicalName,
                ColumnSet = new ColumnSet(Common.Model.ProjectTask.ActualStart, Common.Model.ProjectTask.ActualEnd),
            };
            listofIncompleteChildTasks.Criteria.AddCondition(Common.Model.ProjectTask.ParentTask, ConditionOperator.Equal, parentTaskID);
            listofIncompleteChildTasks.Criteria.AddCondition(Common.Model.ProjectTask.ActualEnd, ConditionOperator.Null);
            listofIncompleteChildTasks.NoLock = true;
            EntityCollection incompleteChildTaskCollection = service.RetrieveMultiple(listofIncompleteChildTasks);
            tracingService.Trace(FormattableString.Invariant($"List of incomplete Task : {incompleteChildTaskCollection.Entities.Count}"));
            if (incompleteChildTaskCollection.Entities.Count < 1)
            {
                QueryExpression listofChildTaskswithEndDate = new QueryExpression
                {
                    EntityName = Common.Model.ProjectTask.LogicalName,
                    ColumnSet = new ColumnSet(Common.Model.ProjectTask.ActualStart, Common.Model.ProjectTask.ActualEnd),
                };
                listofChildTaskswithEndDate.Criteria.AddCondition(Common.Model.ProjectTask.ParentTask, ConditionOperator.Equal, parentTaskID);
                listofChildTaskswithEndDate.Criteria.AddCondition(Common.Model.ProjectTask.ActualEnd, ConditionOperator.NotNull);
                listofChildTaskswithEndDate.NoLock = true;
                EntityCollection endDateChildTaskCollection = service.RetrieveMultiple(listofChildTaskswithEndDate);
                tracingService.Trace(FormattableString.Invariant($"End Date Child Task Count: {endDateChildTaskCollection.Entities.Count}"));
                if (endDateChildTaskCollection.Entities.Any())
                {
                    parentEndDate = endDateChildTaskCollection.Entities.OrderByDescending(endate => endate.Attributes[Common.Model.ProjectTask.ActualEnd]).FirstOrDefault().GetAttributeValue<DateTime>(Common.Model.ProjectTask.ActualEnd);
                    tracingService.Trace("End Date: " + parentEndDate);
                }
            }

            Entity parentTask = service.Retrieve(Common.Model.ProjectTask.LogicalName, parentTaskID, new ColumnSet(Common.Model.ProjectTask.ActualStart, Common.Model.ProjectTask.ActualEnd));
            DateTime? parentTaskExistingStartDate = parentTask.Attributes.ContainsKey(Common.Model.ProjectTask.ActualStart) ? parentTask.GetAttributeValue<DateTime?>(Common.Model.ProjectTask.ActualStart) : null;
            tracingService.Trace("Parent Task Existing Start Date: " + parentTaskExistingStartDate);
            DateTime? parentTaskExistingEndDate = parentTask.Attributes.ContainsKey(Common.Model.ProjectTask.ActualEnd) ? parentTask.GetAttributeValue<DateTime?>(Common.Model.ProjectTask.ActualEnd) : null;
            tracingService.Trace("Parent Task Existing End Date: " + parentTaskExistingEndDate);
            if (parentStartDate != parentTaskExistingStartDate || parentEndDate != parentTaskExistingEndDate)
            {
                projectTaskEntity[Common.Model.ProjectTask.ActualStart] = parentStartDate;
                projectTaskEntity[Common.Model.ProjectTask.ActualEnd] = parentEndDate;
                tracingService.Trace("Project task entity object is ready to be updated");
                service.Update(projectTaskEntity);
                tracingService.Trace("Project Task updated with Actual Start and End Date.");
            }
        }

        /// <summary>
        /// ProjectBPFAutomateStage function.
        /// </summary>
        /// <param name="postImage">post Image.</param>
        /// <param name="tracingService">tracing service.</param>
        /// <param name="targetEntity">target entity.</param>
        /// <param name="service">IOrganization service.</param>
        /// <param name="project">Project entity.</param>
        public static void ProjectBPFAutomateStage(Entity postImage, ITracingService tracingService, Entity targetEntity, IOrganizationService service, EntityReference project)
        {
            bool milestoneFinished = false;
            EntityReference milestone = null;
            if (targetEntity.Contains(Common.Model.ProjectTask.MilestoneFinished) && targetEntity.Attributes[Common.Model.ProjectTask.MilestoneFinished] != null)
            {
                milestoneFinished = (bool)targetEntity.Attributes[Common.Model.ProjectTask.MilestoneFinished];
                if (milestoneFinished == true)
                {
                    tracingService.Trace("Milestone Finished is True");

                    if (postImage != null && postImage.Attributes.Contains(Common.Model.ProjectTask.Milestone) && postImage.Attributes.Contains(Common.Model.ProjectTask.Project) && postImage.Attributes[Common.Model.ProjectTask.Milestone] != null && postImage.Attributes[Common.Model.ProjectTask.Project] != null)
                    {
                        tracingService.Trace("Post Image contains Milestone & Project");
                        milestone = (EntityReference)postImage.Attributes[Common.Model.ProjectTask.Milestone];
                        project = (EntityReference)postImage.Attributes[Common.Model.ProjectTask.Project];
                        ProjectBPFRetrieveValues.RetrieveValues(service, tracingService, project, milestone, "Automate BPF", string.Empty);
                    }
                }
            }
        }

        /// <summary>
        /// Updates work order service task name from project task.
        /// </summary>
        /// <param name="targetEntity">Target entity.</param>
        /// <param name="service">Organization service.</param>
        /// <param name="tracingService">Tracing service.</param>
        public static void UpdateWOSTNameFromProjectTask(Entity targetEntity, IOrganizationService service, ITracingService tracingService)
        {
            if ((targetEntity.Contains(Common.Model.ProjectTask.Subject) && targetEntity.Attributes[Common.Model.ProjectTask.Subject] != null) || (targetEntity.Contains(Common.Model.ProjectTask.EstimatedEffort) && targetEntity.Attributes[Common.Model.ProjectTask.EstimatedEffort] != null))
            {
                tracingService.Trace(FormattableString.Invariant($"Project Task Name has been Updated"));
                EntityCollection workOrderServiceTaskCollection = BusinessLogic.WorkOrderServiceTask.FetchRelatedWorkOrderServiceTask(service, targetEntity.Id, tracingService);
                if (workOrderServiceTaskCollection.Entities.Count == 1)
                {
                    tracingService.Trace(FormattableString.Invariant($"Project Task Contains 1 WOST"));
                    Entity workOrderServiceTask = new Entity(Common.Model.WorkOrderServiceTask.LogicalName);
                    if (workOrderServiceTaskCollection.Entities[0].Contains(Common.Model.WorkOrderServiceTask.WorkOrderServiceTaskId) && workOrderServiceTaskCollection.Entities[0].Attributes[Common.Model.WorkOrderServiceTask.WorkOrderServiceTaskId] != null)
                    {
                        workOrderServiceTask[Common.Model.WorkOrderServiceTask.WorkOrderServiceTaskId] = workOrderServiceTaskCollection.Entities[0].Attributes[Common.Model.WorkOrderServiceTask.WorkOrderServiceTaskId];
                    }

                    if (targetEntity.Contains(Common.Model.ProjectTask.EstimatedEffort) && targetEntity.Attributes[Common.Model.ProjectTask.EstimatedEffort] != null)
                    {
                        workOrderServiceTask[Common.Model.WorkOrderServiceTask.PlannedHours] = (decimal)targetEntity.GetAttributeValue<double>(Common.Model.ProjectTask.EstimatedEffort);
                    }

                    if (targetEntity.Contains(Common.Model.ProjectTask.Subject) && targetEntity.Attributes[Common.Model.ProjectTask.Subject] != null)
                    {
                        workOrderServiceTask[Common.Model.WorkOrderServiceTask.Name] = targetEntity.Attributes[Common.Model.ProjectTask.Subject];
                    }

                    service.Update(workOrderServiceTask);
                    tracingService.Trace("After WorkOrderServiceTask Updation ");
                }
            }
        }

        /// <summary>
        /// To update update WOST booking on change of dates in project task.
        /// </summary>
        /// <param name="targetEntity">Target entity.</param>
        /// <param name="service">Organization service.</param>
        /// <param name="tracingService">Tracing service.</param>
        public static void UpdateWOSTBookingOnChangeofDatesinProjectTask(Entity targetEntity, IOrganizationService service, ITracingService tracingService)
        {
            if ((targetEntity.Contains(Common.Model.ProjectTask.ScheduledStart) && targetEntity.Attributes[Common.Model.ProjectTask.ScheduledStart] != null) || (targetEntity.Contains(Common.Model.ProjectTask.ScheduledEnd) && targetEntity.Attributes[Common.Model.ProjectTask.ScheduledEnd] != null) || (targetEntity.Contains(Common.Model.ProjectTask.InternalNotes) && targetEntity.Attributes[Common.Model.ProjectTask.InternalNotes] != null))
            {
                ProjectTask.UpdateDates(targetEntity.Id, targetEntity, service, tracingService);
            }
        }

        /// <summary>
        /// Create milestone when milestone is completed and is WOST is no.
        /// </summary>
        /// <param name="targetEntity">Target entity.</param>
        /// <param name="postImage">Post image.</param>
        /// <param name="service">Organization service.</param>
        /// <param name="tracingService">Tracing service.</param>
        public static void CreateMilestoneOnMilestoneFinishAndIsWostIsNo(Entity targetEntity, Entity postImage, IOrganizationService service, ITracingService tracingService)
        {
            if (targetEntity.Contains(Common.Model.ProjectTask.MilestoneFinished) && targetEntity.Attributes[Common.Model.ProjectTask.MilestoneFinished] != null)
            {
                bool milestoneFinished = (bool)targetEntity.Attributes[Common.Model.ProjectTask.MilestoneFinished];
                if (milestoneFinished == true)
                {
                    tracingService.Trace(FormattableString.Invariant($"Milestone is Finished"));
                    if (postImage != null && postImage.Attributes.Contains(Common.Model.ProjectTask.IsWoServiceTask) && (bool)postImage.Attributes[Common.Model.ProjectTask.IsWoServiceTask] == false)
                    {
                        tracingService.Trace(FormattableString.Invariant($"IsWOST is false"));
                        Entity milestone = new Entity(Common.Model.Milestone.LogicalName);
                        if (postImage.Attributes.Contains(Common.Model.ProjectTask.Project) && postImage.Attributes[Common.Model.ProjectTask.Project] != null)
                        {
                            Guid projectId = ((EntityReference)postImage.Attributes[Common.Model.ProjectTask.Project]).Id;
                            if (projectId != Guid.Empty)
                            {
                                milestone[Common.Model.Milestone.Project] = new EntityReference(Common.Model.Milestone.LogicalName, projectId);
                            }
                        }

                        if (targetEntity.Id != null)
                        {
                            milestone[Common.Model.Milestone.ProjectTask] = new EntityReference(Common.Model.Milestone.LogicalName, targetEntity.Id);
                        }

                        if (postImage.Attributes.Contains(Common.Model.ProjectTask.ActualStart) && postImage.Attributes[Common.Model.ProjectTask.ActualStart] != null)
                        {
                            DateTime actualStart = (DateTime)postImage.Attributes[Common.Model.ProjectTask.ActualStart];
                            if (actualStart != null)
                            {
                                milestone[Common.Model.Milestone.StartDate] = actualStart;
                            }
                        }

                        if (postImage.Attributes.Contains(Common.Model.ProjectTask.ActualEnd) && postImage.Attributes[Common.Model.ProjectTask.ActualEnd] != null)
                        {
                            DateTime actualEnd = (DateTime)postImage.Attributes[Common.Model.ProjectTask.ActualEnd];
                            if (actualEnd != null)
                            {
                                milestone[Common.Model.Milestone.EndDate] = actualEnd;
                            }
                        }

                        if (postImage.Attributes.Contains(Common.Model.ProjectTask.Milestone) && postImage.Attributes[Common.Model.ProjectTask.Milestone] != null)
                        {
                            Guid milestoneId = ((EntityReference)postImage.Attributes[Common.Model.ProjectTask.Milestone]).Id;
                            string milestoneName = ((EntityReference)postImage.Attributes[Common.Model.ProjectTask.Milestone]).Name;
                            if (milestoneId != Guid.Empty)
                            {
                                milestone[Common.Model.Milestone.MilestoneField] = new EntityReference(Common.Model.Milestone.LogicalName, milestoneId);
                            }

                            if (milestoneName != string.Empty)
                            {
                                milestone[Common.Model.Milestone.Name] = milestoneName;
                            }
                        }

                        milestone[Common.Model.Milestone.MilestoneFinished] = true;
                        service.Create(milestone);
                    }
                }
            }
        }

        /// <summary>
        /// Updates the total of Trouble and Non conformance hours in parent task.
        /// </summary>
        /// <param name="postImage">Post image.</param>
        /// <param name="tracingService">Tracing service.</param>
        /// <param name="targetEntity">Target entity.</param>
        /// <param name="service">Organization service.</param>
        public static void UpdateNonComformanceandTroubleHours(Entity postImage, ITracingService tracingService, Entity targetEntity, IOrganizationService service)
        {
            if (postImage.Contains(Common.Model.ProjectTask.ParentTask) && (targetEntity.Contains(Common.Model.ProjectTask.NonConformance) || targetEntity.Contains(Common.Model.ProjectTask.TroubleHours)))
            {
                tracingService.Trace(FormattableString.Invariant($"UpdateNonComformanceandTroubleHours method is being executed."));
                Guid parentTaskID = ((EntityReference)postImage.Attributes[Common.Model.ProjectTask.ParentTask]).Id;
                tracingService.Trace(FormattableString.Invariant($"Parent Task ID: ") + parentTaskID);
                QueryExpression listofChildTasks = new QueryExpression
                {
                    EntityName = Common.Model.ProjectTask.LogicalName,
                    ColumnSet = new ColumnSet(Common.Model.ProjectTask.NonConformance, Common.Model.ProjectTask.TroubleHours),
                };
                var filter1 = new FilterExpression();
                listofChildTasks.Criteria.AddFilter(filter1);
                filter1.AddCondition(Common.Model.ProjectTask.ParentTask, ConditionOperator.Equal, parentTaskID);
                var filter2 = new FilterExpression();
                filter1.AddFilter(filter2);
                filter2.FilterOperator = LogicalOperator.Or;
                var filter2_1 = new FilterExpression();
                filter2.AddFilter(filter2_1);
                filter2_1.AddCondition(Common.Model.ProjectTask.NonConformance, ConditionOperator.NotNull);
                filter2_1.AddCondition(Common.Model.ProjectTask.NonConformance, ConditionOperator.NotEqual, "0");
                var filter2_2 = new FilterExpression();
                filter2.AddFilter(filter2_2);
                filter2_2.AddCondition(Common.Model.ProjectTask.TroubleHours, ConditionOperator.NotNull);
                filter2_2.AddCondition(Common.Model.ProjectTask.TroubleHours, ConditionOperator.NotEqual, "0");
                listofChildTasks.NoLock = true;
                EntityCollection childTaskCollection = service.RetrieveMultiple(listofChildTasks);
                tracingService.Trace(FormattableString.Invariant($"Child Task Count: {childTaskCollection.Entities.Count}"));
                if (childTaskCollection.Entities.Any())
                {
                    decimal totalnonComformanceHours = 0m;
                    decimal totalTroubleHours = 0m;

                    foreach (Entity pT in childTaskCollection.Entities)
                    {
                        if (pT.Contains(Common.Model.ProjectTask.NonConformance) && pT.Attributes[Common.Model.ProjectTask.NonConformance] != null)
                        {
                            decimal nonComformanceHours = Convert.ToDecimal(pT.Attributes[Common.Model.ProjectTask.NonConformance]);
                            tracingService.Trace(FormattableString.Invariant($"{nonComformanceHours.ToString(CultureInfo.CurrentCulture)}"));

                            if (nonComformanceHours != 0m)
                            {
                                totalnonComformanceHours += nonComformanceHours;
                            }
                        }

                        if (pT.Contains(Common.Model.ProjectTask.TroubleHours) && pT.Attributes[Common.Model.ProjectTask.TroubleHours] != null)
                        {
                            decimal troubleHours = Convert.ToDecimal(pT.Attributes[Common.Model.ProjectTask.TroubleHours]);
                            tracingService.Trace(FormattableString.Invariant($"{troubleHours.ToString(CultureInfo.CurrentCulture)}"));
                            if (troubleHours != 0m)
                            {
                                totalTroubleHours += troubleHours;
                            }
                        }
                    }

                    tracingService.Trace(FormattableString.Invariant($"Before updating NonComformance and Trouble hours in project task"));
                    Entity projectTaskEntity = new Entity(Common.Model.ProjectTask.LogicalName, parentTaskID);
                    projectTaskEntity[Common.Model.ProjectTask.NonConformance] = totalnonComformanceHours;
                    projectTaskEntity[Common.Model.ProjectTask.TroubleHours] = totalTroubleHours;
                    if (totalnonComformanceHours != 0 || totalTroubleHours != 0)
                    {
                        service.Update(projectTaskEntity);
                    }

                    tracingService.Trace(FormattableString.Invariant($"Project Task Updated."));
                }
            }
        }

        /// <summary>
        /// Plugin Execute.
        /// </summary>
        /// <param name="serviceProvider">service provider.</param>
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IPluginExecutionContext4 context4 = (IPluginExecutionContext4)serviceProvider.GetService(typeof(IPluginExecutionContext4));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            Entity targetEntity = null;
            Entity preImage = null;
            Entity postImage = null;
            DateTime previousStartDate = DateTime.MinValue;
            DateTime previousEndDate = DateTime.MinValue;
            DateTime presentStartDate = DateTime.MinValue;
            DateTime presentEndDate = DateTime.MinValue;
            EntityReference project = null;

            // Targets - xmultiple changes
            Entity singleTarget = context.InputParameters.Contains(Common.Model.Common.Target) && context.InputParameters[Common.Model.Common.Target] is Entity ?
    (Entity)context.InputParameters[Common.Model.Common.Target] : null;
            EntityCollection targets = new EntityCollection();
            if (singleTarget == null)
            {
                targets = context.InputParameters.Contains(Common.Model.Common.Targets) && context.InputParameters[Common.Model.Common.Targets] is EntityCollection ?
                    (EntityCollection)context.InputParameters[Common.Model.Common.Targets] : null;
            }
            else
            {
                List<Entity> singleTargetList = new List<Entity>();
                singleTargetList.Add(singleTarget);
                targets = new EntityCollection(singleTargetList);
                tracingService.Trace("Added");
            }

            // PreEntityImages
            List<Entity> preEntityImages = ProjectTask.GetBulkImages(tracingService, context4, context, "PreImage", "PreImage");

            // PostEntityImages
            List<Entity> postEntityImages = ProjectTask.GetBulkImages(tracingService, context4, context, "PostImage", "PostImage");

            // Loop through Targets, PreImages and PostImages
            if (targets != null && targets.Entities != null && targets.Entities.Count > 0)
            {
                for (int i = 0; i < targets.Entities.Count; i++)
                {
                    targetEntity = targets.Entities[i];
                    ////Check if the targetEntity has any of the following field updates:- msdyn_actualend, msdyn_actualstart, msdyn_description, msdyn_effort, msdyn_scheduledend, msdyn_scheduledstart, msdyn_subject, tel_po_comment, tel_po_ignoreformilestone, tel_psa_desc1, tel_psa_desc2, tel_psa_desc3, tel_psa_desc4, tel_psa_desc5, tel_psa_milestone, tel_psa_nonconformance, tel_psa_previousend, tel_psa_previousstart, tel_psa_troublehours, tel_psa_url1, tel_psa_url2, tel_psa_url3, tel_psa_url4, tel_psa_url5
                    if (!(targetEntity.Contains(Common.Model.ProjectTask.ActualStart) && targetEntity.Attributes[Common.Model.ProjectTask.ActualStart] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.ActualEnd) && targetEntity.Attributes[Common.Model.ProjectTask.ActualEnd] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.Description) && targetEntity.Attributes[Common.Model.ProjectTask.Description] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.EstimatedEffort) && targetEntity.Attributes[Common.Model.ProjectTask.EstimatedEffort] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.ScheduledEnd) && targetEntity.Attributes[Common.Model.ProjectTask.ScheduledEnd] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.ScheduledStart) && targetEntity.Attributes[Common.Model.ProjectTask.ScheduledStart] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.Subject) && targetEntity.Attributes[Common.Model.ProjectTask.Subject] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.Comments) && targetEntity.Attributes[Common.Model.ProjectTask.Comments] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.IgnoreForMilestone) && targetEntity.Attributes[Common.Model.ProjectTask.IgnoreForMilestone] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.Desc1) && targetEntity.Attributes[Common.Model.ProjectTask.Desc1] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.Desc2) && targetEntity.Attributes[Common.Model.ProjectTask.Desc2] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.Desc3) && targetEntity.Attributes[Common.Model.ProjectTask.Desc3] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.Desc4) && targetEntity.Attributes[Common.Model.ProjectTask.Desc4] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.Desc5) && targetEntity.Attributes[Common.Model.ProjectTask.Desc5] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.Milestone) && targetEntity.Attributes[Common.Model.ProjectTask.Milestone] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.NonConformance) && targetEntity.Attributes[Common.Model.ProjectTask.NonConformance] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.PreviousEnd) && targetEntity.Attributes[Common.Model.ProjectTask.PreviousEnd] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.PreviousStart) && targetEntity.Attributes[Common.Model.ProjectTask.PreviousStart] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.TroubleHours) && targetEntity.Attributes[Common.Model.ProjectTask.TroubleHours] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.Url1) && targetEntity.Attributes[Common.Model.ProjectTask.Url1] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.Url2) && targetEntity.Attributes[Common.Model.ProjectTask.Url2] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.Url3) && targetEntity.Attributes[Common.Model.ProjectTask.Url3] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.Url4) && targetEntity.Attributes[Common.Model.ProjectTask.Url4] != null) && !(targetEntity.Contains(Common.Model.ProjectTask.Url5) && targetEntity.Attributes[Common.Model.ProjectTask.Url5] != null))
                    {
                        continue;
                    }

                    preImage = preEntityImages != null && preEntityImages.Any() ? preEntityImages[i] : null;
                    postImage = postEntityImages != null && postEntityImages.Any() ? postEntityImages[i] : null;

                    tracingService.Trace("Target Entity: " + targetEntity.Id.ToString());

                    tracingService.Trace(FormattableString.Invariant($"Depth: {context.Depth}"));

                    if ((targetEntity.Contains(Common.Model.ProjectTask.ScheduledStart) && (targetEntity.Attributes[Common.Model.ProjectTask.ScheduledStart] != null)) ||
                         (targetEntity.Contains(Common.Model.ProjectTask.ScheduledEnd) && (targetEntity.Attributes[Common.Model.ProjectTask.ScheduledEnd] != null)) || (targetEntity.Contains(Common.Model.ProjectTask.InternalNotes) && (targetEntity.Attributes[Common.Model.ProjectTask.InternalNotes] != null)))
                    {
                        UpdateWOSTBookingOnChangeofDatesinProjectTask(targetEntity, service, tracingService);
                    }

                    if (context.Depth < 3)
                    {
                        if (context.PreEntityImages.Contains(Common.Model.Common.PreImage))
                        {
                            preImage = context.PreEntityImages[Common.Model.Common.PreImage];
                            if (preImage.Attributes.Contains(Common.Model.ProjectTask.ScheduledStart) && preImage.Attributes[Common.Model.ProjectTask.ScheduledStart] != null && preImage.Attributes.Contains(Common.Model.ProjectTask.ScheduledEnd) && preImage.Attributes[Common.Model.ProjectTask.ScheduledEnd] != null)
                            {
                                previousStartDate = (DateTime)preImage.Attributes[Common.Model.ProjectTask.ScheduledStart];
                                previousEndDate = (DateTime)preImage.Attributes[Common.Model.ProjectTask.ScheduledEnd];
                                project = (EntityReference)preImage.Attributes[Common.Model.ProjectTask.Project];
                                tracingService.Trace(FormattableString.Invariant($"Previous Start  {previousStartDate.ToString(CultureInfo.InvariantCulture)}, Previous End {previousEndDate.ToString(CultureInfo.InvariantCulture)}"));
                            }
                        }

                        if ((postImage.Contains(Common.Model.ProjectTask.ParentTask)
                            && (targetEntity.Contains(Common.Model.ProjectTask.NonConformance)
                            && targetEntity.Attributes[Common.Model.ProjectTask.NonConformance] != null)) || (targetEntity.Contains(Common.Model.ProjectTask.TroubleHours)
                            && targetEntity.Attributes[Common.Model.ProjectTask.TroubleHours] != null))
                        {
                            UpdateNonComformanceandTroubleHours(postImage, tracingService, targetEntity, service);
                        }

                        if (postImage.Contains(Common.Model.ProjectTask.ParentTask) && (targetEntity.Contains(Common.Model.ProjectTask.ActualStart) || targetEntity.Contains(Common.Model.ProjectTask.ActualEnd)))
                        {
                            UpdateActualStartandActualEndDate(postImage, tracingService, targetEntity, service);
                        }

                        if (postImage.Contains(Common.Model.ProjectTask.ParentTask) && ((targetEntity.Contains(Common.Model.ProjectTask.ActualStart) && targetEntity.Attributes[Common.Model.ProjectTask.ActualStart] != null) || (targetEntity.Contains(Common.Model.ProjectTask.ActualEnd) && targetEntity.Attributes[Common.Model.ProjectTask.ActualEnd] != null) || (targetEntity.Contains(Common.Model.ProjectTask.IgnoreForMilestone) && targetEntity.Attributes[Common.Model.ProjectTask.IgnoreForMilestone] != null)))
                        {
                            this.UpdateMilestoneFieldsOnParent(postImage, tracingService, targetEntity, service, context.InitiatingUserId);
                        }

                        if (postImage.Contains(Common.Model.ProjectTask.OutlineLevel) && postImage.Contains(Common.Model.ProjectTask.Project) && postImage[Common.Model.ProjectTask.Project] != null && ((targetEntity.Contains(Common.Model.ProjectTask.ActualStart) && targetEntity.Attributes[Common.Model.ProjectTask.ActualStart] != null) || (targetEntity.Contains(Common.Model.ProjectTask.ActualEnd) && targetEntity.Attributes[Common.Model.ProjectTask.ActualEnd] != null)))
                        {
                            tracingService.Trace("Checking 'outline equals to 1' condition to execute UpdateProjectActualStartandActualEndDate");
                            int outlineLevel1 = postImage.GetAttributeValue<int>(Common.Model.ProjectTask.OutlineLevel);
                            tracingService.Trace(FormattableString.Invariant($"Outline Level : {outlineLevel1}"));
                            if (outlineLevel1 == 1)
                            {
                                this.UpdateProjectActualStartandActualEndDate(postImage, targetEntity, tracingService, service);
                            }
                        }

                        if (context.Depth == 1)
                        {
                            ProjectBPFAutomateStage(postImage, tracingService, targetEntity, service, project);
                            //// CreateMilestoneOnMilestoneFinishAndIsWostIsNo(targetEntity, postImage, service, tracingService);
                        }
                        else if (context.Depth == 2)
                        {
                            EnabledForSnapshot(targetEntity, service, tracingService, preImage, postImage, presentStartDate, presentEndDate, previousStartDate, previousEndDate, project);

                            ////[Commented as part of #52622, Effort and Subject name update is added in the power automate 15. Psa.Flow.ProjectTask.copy Milestone Field & Comments from PT to WOST(Update)]
                            //// UpdateWOSTNameFromProjectTask(targetEntity, service, tracingService);
                        }
                    }
                    else
                    {
                        if (context.Depth == 3 && postImage.Contains(Common.Model.ProjectTask.OutlineLevel) && postImage.Contains(Common.Model.ProjectTask.Project) && postImage[Common.Model.ProjectTask.Project] != null && ((targetEntity.Contains(Common.Model.ProjectTask.ActualStart) && targetEntity.Attributes[Common.Model.ProjectTask.ActualStart] != null) || (targetEntity.Contains(Common.Model.ProjectTask.ActualEnd) && targetEntity.Attributes[Common.Model.ProjectTask.ActualEnd] != null)))
                        {
                            tracingService.Trace("Checking 'outline equals to 1' condition to execute UpdateProjectActualStartandActualEndDate");
                            int outlineLevel = postImage.GetAttributeValue<int>(Common.Model.ProjectTask.OutlineLevel);
                            tracingService.Trace(FormattableString.Invariant($"Outline Level: {outlineLevel}"));
                            if (outlineLevel == 1)
                            {
                                this.UpdateProjectActualStartandActualEndDate(postImage, targetEntity, tracingService, service);
                            }
                        }
                        else
                        {
                            tracingService.Trace("Returning as depth is greater than 2");
                            if (context.ParentContext != null)
                            {
                                tracingService.Trace("Parent Context Message:" + context.ParentContext.MessageName);
                                tracingService.Trace("Initiating User:" + context.ParentContext.InitiatingUserId);
                            }

                         
                        }
                    }
                }
            }
        }

        /// <summary>
        /// To UpdateMilestoneFieldsOnParent.
        /// </summary>
        /// <param name="postImage">For postImage.</param>
        /// <param name="tracingService">Tracing Service.</param>
        /// <param name="targetEntity">Target Entity.</param>
        /// <param name="service">For Service.</param>
        /// <param name="userId">User ID.</param>
        public void UpdateMilestoneFieldsOnParent(Entity postImage, ITracingService tracingService, Entity targetEntity, IOrganizationService service, Guid userId)
        {
            Guid parentTaskID = ((EntityReference)postImage.Attributes[Common.Model.ProjectTask.ParentTask]).Id;
            ProjectTask.UpdateParentTaskMilestoneFields(parentTaskID, targetEntity, tracingService, service, userId);
        }

        /// <summary>
        /// This method updates the start date and end date.
        /// </summary>
        /// <param name="postImage">Post image.</param>
        /// <param name="targetEntity">target entity.</param>
        /// <param name="tracingService">tracing service.</param>
        /// <param name="service">organization service.</param>
        private void UpdateProjectActualStartandActualEndDate(Entity postImage, Entity targetEntity, ITracingService tracingService, IOrganizationService service)
        {
            tracingService.Trace("UpdateProjectActualStartandActualEndDate method is being executed.");
            DateTime? actualStartDate = null;
            DateTime? actualEndDate = null;
            Guid projectID = ((EntityReference)postImage[Common.Model.ProjectTask.Project]).Id;
            QueryExpression queryExpression = new QueryExpression()
            {
                EntityName = Common.Model.ProjectTask.LogicalName,
                ColumnSet = new ColumnSet(Common.Model.ProjectTask.ActualStart, Common.Model.ProjectTask.ActualEnd),
            };
            queryExpression.Criteria.AddCondition(Common.Model.ProjectTask.OutlineLevel, ConditionOperator.Equal, 1);
            queryExpression.Criteria.AddCondition(Common.Model.ProjectTask.Project, ConditionOperator.Equal, projectID);
            if (targetEntity.Contains(Common.Model.ProjectTask.ActualStart))
            {
                queryExpression.Criteria.AddCondition(Common.Model.ProjectTask.ActualStart, ConditionOperator.NotNull);
                queryExpression.NoLock = true;
                EntityCollection projectTasksStartDate = service.RetrieveMultiple(queryExpression);
                tracingService.Trace(FormattableString.Invariant($"Start Date Layer1 Task Count: {projectTasksStartDate.Entities.Count}"));
                if (projectTasksStartDate.Entities.Any())
                {
                    actualStartDate = projectTasksStartDate.Entities.OrderBy(stdate => stdate.Attributes[Common.Model.ProjectTask.ActualStart]).FirstOrDefault().GetAttributeValue<DateTime>(Common.Model.ProjectTask.ActualStart);
                    tracingService.Trace("Project Start Date: " + actualStartDate);
                }
            }

            if (targetEntity.Contains(Common.Model.ProjectTask.ActualEnd))
            {
                QueryExpression queryExpressionIncompleteTask = new QueryExpression()
                {
                    EntityName = Common.Model.ProjectTask.LogicalName,
                    ColumnSet = new ColumnSet(Common.Model.ProjectTask.ActualStart, Common.Model.ProjectTask.ActualEnd),
                };
                queryExpressionIncompleteTask.Criteria.AddCondition(Common.Model.ProjectTask.OutlineLevel, ConditionOperator.Equal, 1);
                queryExpressionIncompleteTask.Criteria.AddCondition(Common.Model.ProjectTask.Project, ConditionOperator.Equal, projectID);
                queryExpressionIncompleteTask.Criteria.AddCondition(Common.Model.ProjectTask.ActualEnd, ConditionOperator.Null);
                queryExpressionIncompleteTask.NoLock = true;
                EntityCollection incompleteProjectTasks = service.RetrieveMultiple(queryExpressionIncompleteTask);
                tracingService.Trace(FormattableString.Invariant($"Incomplete layer 1 task count: {incompleteProjectTasks.Entities.Count}"));
                if (incompleteProjectTasks.Entities.Count < 1)
                {
                    QueryExpression queryExpressionEndDate = new QueryExpression()
                    {
                        EntityName = Common.Model.ProjectTask.LogicalName,
                        ColumnSet = new ColumnSet(Common.Model.ProjectTask.ActualStart, Common.Model.ProjectTask.ActualEnd),
                    };
                    queryExpressionEndDate.Criteria.AddCondition(Common.Model.ProjectTask.OutlineLevel, ConditionOperator.Equal, 1);
                    queryExpressionEndDate.Criteria.AddCondition(Common.Model.ProjectTask.Project, ConditionOperator.Equal, projectID);
                    queryExpressionEndDate.Criteria.AddCondition(Common.Model.ProjectTask.ActualEnd, ConditionOperator.NotNull);
                    queryExpressionEndDate.NoLock = true;
                    EntityCollection projectTasksEndDate = service.RetrieveMultiple(queryExpressionEndDate);
                    tracingService.Trace(FormattableString.Invariant($"End Date Layer1 Task Count: {projectTasksEndDate.Entities.Count}"));
                    if (projectTasksEndDate.Entities.Any())
                    {
                        actualEndDate = projectTasksEndDate.Entities.OrderByDescending(endate => endate.Attributes[Common.Model.ProjectTask.ActualEnd]).FirstOrDefault().GetAttributeValue<DateTime>(Common.Model.ProjectTask.ActualEnd);
                        tracingService.Trace(" Project End Date: " + actualEndDate);
                    }
                }
            }

            Entity project = new Entity(Common.Model.Project.LogicalName, projectID);
            project[Common.Model.Project.ActualStartDate] = actualStartDate;
            project[Common.Model.Project.ActualEndDate] = actualEndDate;
            tracingService.Trace("Before update start date: " + project.GetAttributeValue<DateTime?>(Common.Model.Project.ActualStartDate));
            tracingService.Trace("Before update start date: " + project.GetAttributeValue<DateTime?>(Common.Model.Project.ActualEndDate));
            tracingService.Trace("Updating the Project with actual dates.");
            service.Update(project);
            tracingService.Trace("Project is updated with actual start date and actual end date.");
        }
    }
}
