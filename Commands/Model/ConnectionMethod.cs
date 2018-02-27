﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.PowerShell.Commands.Model
{
    public enum ConnectionMethod
    {
        WebLogin,
        Credentials,
        AccessToken,
        AzureADAppOnly,
        AzureADNativeApplication,
        ADFS,
        GraphDeviceLogin
    }
}
