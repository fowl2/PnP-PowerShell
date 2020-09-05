using System.Management.Automation;

using Microsoft.SharePoint.Client;

using PnP.PowerShell.CmdletHelpAttributes;
using PnP.PowerShell.Commands.Base.PipeBinds;

namespace PnP.PowerShell.Commands.ContentTypes
{
    [Cmdlet(VerbsCommon.Add, "PnPContentTypeToList")]
    [CmdletHelp("Adds a new content type to a list",
        Category = CmdletHelpCategory.ContentTypes)]
    [CmdletExample(
        Code = @"PS:> Add-PnPContentTypeToList -List ""Documents"" -ContentType ""Project Document"" -DefaultContentType",
        Remarks = @"This will add an existing content type to a list and sets it as the default content type",
        SortOrder = 1)]
    public class AddContentTypeToList : PnPWebCmdlet
    {
        [Parameter(Mandatory = true, HelpMessage = "Specifies the list to which the content type needs to be added")]
        [ValidateNotNullOrEmpty]
        public ListPipeBind List;

        [Parameter(Mandatory = true, HelpMessage = "Specifies the content type that needs to be added to the list")]
        [ValidateNotNullOrEmpty]
        public ContentTypePipeBind ContentType;

        [Parameter(Mandatory = false, HelpMessage = "Specify if the content type needs to be the default content type or not")]
        public SwitchParameter DefaultContentType;

        protected override void ExecuteCmdlet()
        {
            var list = List.GetListOrThrow(nameof(List), SelectedWeb);
            var ct = ContentType?.GetContentTypeOrWarn(this, list);

            if (ct != null)
            {
                SelectedWeb.AddContentTypeToList(list.Title, ct, DefaultContentType);
            }
        }
    }
}
