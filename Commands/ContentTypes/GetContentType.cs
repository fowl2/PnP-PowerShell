using System.Linq;
using System.Management.Automation;
using Microsoft.SharePoint.Client;
using PnP.PowerShell.CmdletHelpAttributes;
using PnP.PowerShell.Commands.Base;
using PnP.PowerShell.Commands.Base.PipeBinds;
using System;

namespace PnP.PowerShell.Commands.ContentTypes
{
    [Cmdlet(VerbsCommon.Get, "PnPContentType")]
    [CmdletHelp("Retrieves a content type",
        Category = CmdletHelpCategory.ContentTypes,
        OutputType = typeof(ContentType),
        OutputTypeLink = "https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.contenttype.aspx")]
    [CmdletExample(
        Code = @"PS:> Get-PnPContentType ",
        Remarks = @"This will get a listing of all available content types within the current web",
        SortOrder = 1)]
    [CmdletExample(
        Code = @"PS:> Get-PnPContentType -InSiteHierarchy",
        Remarks = @"This will get a listing of all available content types within the site collection",
        SortOrder = 2)]
    [CmdletExample(
        Code = @"PS:> Get-PnPContentType -Identity ""Project Document""",
        Remarks = @"This will get the content type with the name ""Project Document"" within the current context",
        SortOrder = 3)]
    [CmdletExample(
        Code = @"PS:> Get-PnPContentType -List ""Documents""",
        Remarks = @"This will get a listing of all available content types within the list ""Documents""",
        SortOrder = 4)]
    public class GetContentType : PnPWebCmdlet
    {
        [Parameter(Mandatory = false, Position = 0, ValueFromPipeline = true, HelpMessage = "Name or ID of the content type to retrieve")]
        [ValidateNotNullOrEmpty]
        public ContentTypePipeBind Identity;

        [Parameter(Mandatory = false, ValueFromPipeline = true, HelpMessage = "List to query")]
        [ValidateNotNullOrEmpty]
        public ListPipeBind List;

        [Parameter(Mandatory = false, ValueFromPipeline = false, HelpMessage = "Search site hierarchy for content types")]
        public SwitchParameter InSiteHierarchy;

        protected override void ExecuteCmdlet()
        {
            if (List != null)
            {
                var list = List?.GetListOrThrow(nameof(List), SelectedWeb);

                if (Identity != null)
                {
                    var ct = Identity.GetContentTypeOrError(this, nameof(Identity), list);

                    if (ct is null)
                        return;

                    WriteObject(ct, false);
                }
                else
                {
                    var cts = ClientContext.LoadQuery(list.ContentTypes.Include(ct => ct.Id, ct => ct.Name, ct => ct.StringId, ct => ct.Group));
                    ClientContext.ExecuteQueryRetry();
                    WriteObject(cts, true);
                }
            }
            else
            {
                if (Identity != null)
                {
                    var ct = Identity.GetContentTypeOrError(this, nameof(Identity), SelectedWeb, InSiteHierarchy);

                    if (ct is null)
                        return;

                    WriteObject(ct, false);
                }
                else
                {
                    var cts = InSiteHierarchy
                    ? ClientContext.LoadQuery(SelectedWeb.AvailableContentTypes)
                    : ClientContext.LoadQuery(SelectedWeb.ContentTypes);

                    ClientContext.ExecuteQueryRetry();

                    WriteObject(cts, true);
                }
            }
        }
    }
}

