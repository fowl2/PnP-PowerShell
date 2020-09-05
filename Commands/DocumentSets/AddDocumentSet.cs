using System.Management.Automation;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.DocumentSet;
using PnP.PowerShell.CmdletHelpAttributes;
using PnP.PowerShell.Commands.Base.PipeBinds;
using System.Linq;

namespace PnP.PowerShell.Commands.DocumentSets
{
    [Cmdlet(VerbsCommon.Add, "PnPDocumentSet")]
    [CmdletHelp("Creates a new document set in a library.",
      Category = CmdletHelpCategory.DocumentSets,
        OutputType = typeof(string))]
    [CmdletExample(
      Code = @"PS:> Add-PnPDocumentSet -List ""Documents"" -ContentType ""Test Document Set"" -Name ""Test""",
      Remarks = "This will add a new document set based upon the 'Test Document Set' content type to a list called 'Documents'. The document set will be named 'Test'",
      SortOrder = 1)]
    public class AddDocumentSet : PnPWebCmdlet
    {
        [Parameter(Mandatory = true, HelpMessage = "The name of the list, its ID or an actual list object from where the document set needs to be added")]
        [ValidateNotNullOrEmpty]
        public ListPipeBind List;

        [Parameter(Mandatory = true, HelpMessage = "The name of the document set")]
        [ValidateNotNullOrEmpty]
        public string Name;

        [Parameter(Mandatory = true, HelpMessage = "The name of the content type, its ID or an actual content object referencing to the document set")]
        [ValidateNotNullOrEmpty]
        public ContentTypePipeBind ContentType;

        protected override void ExecuteCmdlet()
        {
            var list = List.GetListOrThrow(nameof(List), SelectedWeb,
                l => l.RootFolder, l => l.ContentTypes);

            var listContentType = ContentType.GetContentType(list);
            if (listContentType is null)
            {
                var siteContentType = ContentType.GetContentTypeOrThrow(nameof(ContentType), SelectedWeb);

                listContentType = new ContentTypePipeBind(siteContentType.Name)
                    .GetContentTypeOrThrow(nameof(ContentType), list);
            }

            if (!listContentType.StringId.StartsWith("0x0120D520"))
                throw new PSArgumentException($"Content type '{ContentType}' does not inherit from the base DocumentSet content type. DocumentSet content type IDs start with 0x0120D520.", nameof(ContentType));

            // Create the document set
            var result = DocumentSet.Create(ClientContext, list.RootFolder, Name, listContentType.Id);
            ClientContext.ExecuteQueryRetry();

            WriteObject(result.Value);
        }
    }
}