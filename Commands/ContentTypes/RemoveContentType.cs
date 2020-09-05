using System.Management.Automation;

using Microsoft.SharePoint.Client;

using PnP.PowerShell.CmdletHelpAttributes;
using PnP.PowerShell.Commands.Base.PipeBinds;

using Resources = PnP.PowerShell.Commands.Properties.Resources;

namespace PnP.PowerShell.Commands.ContentTypes
{
    [Cmdlet(VerbsCommon.Remove, "PnPContentType")]
    [CmdletHelp("Removes a content type from a web",
        Category = CmdletHelpCategory.ContentTypes)]
    [CmdletExample(
        Code = @"PS:> Remove-PnPContentType -Identity ""Project Document""",
        Remarks = @"This will remove a content type called ""Project Document"" from the current web",
        SortOrder = 1)]
    [CmdletExample(
        Code = @"PS:> Remove-PnPContentType -Identity ""Project Document"" -Force",
        Remarks = @"This will remove a content type called ""Project Document"" from the current web with force",
        SortOrder = 2)]
    public class RemoveContentType : PnPWebCmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, HelpMessage = "The name or ID of the content type to remove")]
        public ContentTypePipeBind Identity;

        [Parameter(Mandatory = false, HelpMessage = "Specifying the Force parameter will skip the confirmation question.")]
        public SwitchParameter Force;

        protected override void ExecuteCmdlet()
        {
            var ct = Identity?.GetContentTypeOrWarn(this, SelectedWeb);

            if (ct != null || Force || ShouldContinue(Resources.RemoveContentType, Resources.Confirm))
            {
                ct.DeleteObject();
                ClientContext.ExecuteQueryRetry();
            }
        }
    }
}
