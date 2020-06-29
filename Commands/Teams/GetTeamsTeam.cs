﻿using SharePointPnP.PowerShell.CmdletHelpAttributes;
using SharePointPnP.PowerShell.Commands.Base;
using SharePointPnP.PowerShell.Commands.Base.PipeBinds;
using SharePointPnP.PowerShell.Commands.Model.Teams;
using SharePointPnP.PowerShell.Commands.Utilities;
using System.Management.Automation;

namespace SharePointPnP.PowerShell.Commands.Graph
{
    [Cmdlet(VerbsCommon.Get, "PnPTeamsTeam")]
    [CmdletHelp("Gets one Microsoft Teams Team or a list of Teams.",
        Category = CmdletHelpCategory.Teams,
        SupportedPlatform = CmdletSupportedPlatform.Online)]
    [CmdletExample(
       Code = "PS:> Get-PnPTeamsTeam",
       Remarks = "Retrieves all the Microsoft Teams instances",
       SortOrder = 1)]
    [CmdletExample(
       Code = "PS:> Get-PnPTeamsTeam -GroupId $groupId",
       Remarks = "Retrieves a specific Microsoft Teams instance",
       SortOrder = 2)]
    [CmdletExample(
       Code = "PS:> Get-PnPTeamsTeam -Visibility Public",
       Remarks = "Retrieves all Microsoft Teams instances which are public visible",
       SortOrder = 2)]
    [CmdletMicrosoftGraphApiPermission(MicrosoftGraphApiPermission.Group_Read_All)]
    [CmdletMicrosoftGraphApiPermission(MicrosoftGraphApiPermission.Group_ReadWrite_All)]
    public class GetTeamsTeam : PnPGraphCmdlet
    {
        private const string ParameterSet_GroupId = "Retrieve a specific Team";

        [Parameter(Mandatory = false, HelpMessage = "Specify the group id of the team to retrieve.")]
        public GuidPipeBind GroupId;

        [Parameter(Mandatory = false, HelpMessage = "Specify the visibility of the teams to retrieve.")]
        public GroupVisibility Visibility;

        protected override void ExecuteCmdlet()
        {
            if (ParameterSpecified(nameof(GroupId)))
            {
                WriteObject(TeamsUtility.GetTeam(AccessToken, HttpClient, GroupId.Id.ToString()));
            }
            else if (ParameterSpecified(nameof(Visibility)))
            {
                WriteObject(TeamsUtility.GetTeams(AccessToken, HttpClient, Visibility), true);
            }
            else
            {
                WriteObject(TeamsUtility.GetTeams(AccessToken, HttpClient), true);
            }
        }
    }
}
