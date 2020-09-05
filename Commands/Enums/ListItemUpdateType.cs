using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PnP.PowerShell.Commands.Enums
{
    public enum ListItemUpdateType
    {
        /// <summary>
        /// Creates a new version and updates "Modified" and "Modified by"
        /// </summary>
        Update,
#if !ONPREMISES
        /// <summary>
        /// Does not create a new version or update "Modified" and "Modified by"
        /// </summary>
        SystemUpdate,
        /// <summary>
        /// Does not create a new version but does update "Modified" and "Modified by"
        /// </summary>
        UpdateOverwriteVersion
#endif
    }
}
