using System.Management.Automation;

using Microsoft.SharePoint.Client;

using PnP.PowerShell.CmdletHelpAttributes;
using PnP.PowerShell.Commands.Base.PipeBinds;

namespace PnP.PowerShell.Commands.ContentTypes
{

    [Cmdlet(VerbsCommon.Remove, "PnPContentTypeFromList")]
    [CmdletHelp("Removes a content type from a list",
        Category = CmdletHelpCategory.ContentTypes)]
    [CmdletExample(
        Code = @"PS:> Remove-PnPContentTypeFromList -List ""Documents"" -ContentType ""Project Document""",
        Remarks = @"This will remove a content type called ""Project Document"" from the ""Documents"" list",
        SortOrder = 1)]
    public class RemoveContentTypeFromList : PnPWebCmdlet
    {
        [Parameter(Mandatory = true, HelpMessage = "The name of the list, its ID or an actual list object from where the content type needs to be removed from")]
        [ValidateNotNullOrEmpty]
        public ListPipeBind List;

        [Parameter(Mandatory = true, HelpMessage = "The name of a content type, its ID or an actual content type object that needs to be removed from the specified list.")]
        [ValidateNotNullOrEmpty]
        public ContentTypePipeBind ContentType;

        protected override void ExecuteCmdlet()
        {
            var list = List.GetListOrThrow(nameof(List), SelectedWeb);
            var ct = ContentType.GetContentTypeOrWarn(this, list);
            if (ct != null)
            {
                SelectedWeb.RemoveContentTypeFromList(list, ct);
            }
        }
    }
}
