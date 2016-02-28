# Start-FHBackup.ps1
# Defines the Start-FHBackup cmdlet.

#region LICENSE
<#

# The MIT License (MIT)
#
# Copyright (c) 2016 mchubby https://github.com/mchubby
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights to
# use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
# of the Software, and to permit persons to whom the Software is furnished to do
# so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

#>
#endregion

Add-Type -TypeDef @"
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;


namespace Mchubby.Interop
{
  using FH_SERVICE_PIPE_HANDLE = System.IntPtr;


  public enum _FH_PROTECTED_ITEM_CATEGORY
  {
    FH_FOLDER,
    FH_LIBRARY,
    MAX_PROTECTED_ITEM_CATEGORY,
  }
  public enum _FH_LOCAL_POLICY_TYPE
  {
    FH_FREQUENCY,
    FH_RETENTION_TYPE,
    FH_RETENTION_AGE,
    MAX_LOCAL_POLICY,
  }
  public enum _FH_BACKUP_STATUS
  {
    FH_STATUS_DISABLED,
    FH_STATUS_DISABLED_BY_GP,
    FH_STATUS_ENABLED,
    FH_STATUS_REHYDRATING,
    MAX_BACKUP_STATUS,
  }
  public enum _FH_TARGET_PROPERTY_TYPE
  {
    FH_TARGET_NAME,
    FH_TARGET_URL,
    FH_TARGET_DRIVE_TYPE,
    MAX_TARGET_PROPERTY,
  }
  public enum _FH_DEVICE_VALIDATION_RESULT
  {
    FH_ACCESS_DENIED,
    FH_INVALID_DRIVE_TYPE,
    FH_READ_ONLY_PERMISSION,
    FH_CURRENT_DEFAULT,
    FH_NAMESPACE_EXISTS,
    FH_TARGET_PART_OF_LIBRARY,
    FH_VALID_TARGET,
    MAX_VALIDATION_RESULT,
  }

  [ComImport,
  InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
  Guid("D87965FD-2BAD-4657-BD3B-9567EB300CED")]
  public interface IFhTarget
  {
    void GetStringProperty([In] _FH_TARGET_PROPERTY_TYPE PropertyType, [MarshalAs(UnmanagedType.BStr)] out string PropertyValue);
    void GetNumericalProperty([In] _FH_TARGET_PROPERTY_TYPE PropertyType, out ulong PropertyValue);
  }

  [ComImport,
  InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
  Guid("3197ABCE-532A-44C6-8615-F3666566A720")]
  public interface IFhScopeIterator
  {
    void MoveToNextItem();
    void GetItem([MarshalAs(UnmanagedType.BStr)] out string Item);
  }

  [ComImport,
  InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
  Guid("6A5FEA5B-BF8F-4EE5-B8C3-44D8A0D7331C")]
  public interface IFhConfigMgr
  {
    void LoadConfiguration();
    void CreateDefaultConfiguration([In] int OverwriteIfExists);
    void SaveConfiguration();
    void AddRemoveExcludeRule([In] int Add, [In] _FH_PROTECTED_ITEM_CATEGORY Category, [MarshalAs(UnmanagedType.BStr), In] string Item);
    void GetIncludeExcludeRules([In] int Include, [In] _FH_PROTECTED_ITEM_CATEGORY Category, [MarshalAs(UnmanagedType.Interface)] out IFhScopeIterator Iterator);
    void GetLocalPolicy([In] _FH_LOCAL_POLICY_TYPE LocalPolicyType, out ulong PolicyValue);
    void SetLocalPolicy([In] _FH_LOCAL_POLICY_TYPE LocalPolicyType, [In] ulong PolicyValue);
    void GetBackupStatus(out _FH_BACKUP_STATUS BackupStatus);
    void SetBackupStatus([In] _FH_BACKUP_STATUS BackupStatus);
    void GetDefaultTarget([MarshalAs(UnmanagedType.Interface)] out IFhTarget DefaultTarget);
    void ValidateTarget([MarshalAs(UnmanagedType.BStr), In] string TargetUrl, out _FH_DEVICE_VALIDATION_RESULT ValidationResult);
    void ProvisionAndSetNewTarget([MarshalAs(UnmanagedType.BStr), In] string TargetUrl, [MarshalAs(UnmanagedType.BStr), In] string TargetName);
    void ChangeDefaultTargetRecommendation([In] int Recommend);
    void QueryProtectionStatus(out uint ProtectionState, [MarshalAs(UnmanagedType.BStr)] out string ProtectedUntilTime);
  }

  public class NativeConstants {
    public static readonly IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);
    public const uint FHSVC_E_BACKUP_BLOCKED = 0x80040600;
    public const uint FHSVC_E_NOT_CONFIGURED = 0x80040601;
    public const uint FHSVC_E_CONFIG_DISABLED = 0x80040602;
    public const uint FHSVC_E_CONFIG_DISABLED_GP = 0x80040603;
    public const uint FHSVC_E_FATAL_CONFIG_ERROR = 0x80040604;
    public const uint FHSVC_E_CONFIG_REHYDRATING = 0x80040605;
  }

  public class NativeMethods {

    // all return values are HRESULTs (see NativeConstants).

    [DllImport("fhsvcctl.dll", PreserveSig = false)]
    public static extern  uint FhServiceOpenPipe([In, MarshalAs(UnmanagedType.Bool)] bool StartServiceIfStopped, out FH_SERVICE_PIPE_HANDLE Pipe);

    [DllImport("fhsvcctl.dll", PreserveSig = false)]
    public static extern  uint FhServiceClosePipe(FH_SERVICE_PIPE_HANDLE Pipe) ;

    [DllImport("fhsvcctl.dll", PreserveSig = false)]
    public static extern  uint FhServiceStartBackup(FH_SERVICE_PIPE_HANDLE Pipe, [In, MarshalAs(UnmanagedType.Bool)] bool LowPriorityIo) ;

    [DllImport("fhsvcctl.dll", PreserveSig = false)]
    public static extern  uint FhServiceStopBackup(FH_SERVICE_PIPE_HANDLE Pipe, [In, MarshalAs(UnmanagedType.Bool)] bool StopTracking) ;

    [DllImport("fhsvcctl.dll", PreserveSig = false)]
    public static extern  uint FhServiceReloadConfiguration(FH_SERVICE_PIPE_HANDLE Pipe) ;

    [DllImport("fhsvcctl.dll", PreserveSig = false)]
    public static extern  uint FhServiceBlockBackup(FH_SERVICE_PIPE_HANDLE Pipe) ;

    [DllImport("fhsvcctl.dll", PreserveSig = false)]
    public static extern  uint FhServiceUnblockBackup(FH_SERVICE_PIPE_HANDLE Pipe) ;

  }


  public class FhConfigMgr
  {
    private IFhConfigMgr _FhConfigMgr;

    public FhConfigMgr()
    {
        _FhConfigMgr = Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("ED43BB3C-09E9-498a-9DF6-2177244C6DB4"))) as IFhConfigMgr;
        _FhConfigMgr.LoadConfiguration();
    }

    public _FH_BACKUP_STATUS BackupStatus
    {
        get
        {
            _FH_BACKUP_STATUS backupStatus;
            _FhConfigMgr.GetBackupStatus(out backupStatus);
            return backupStatus;
        }
        set
        {
            _FhConfigMgr.SetBackupStatus(value);
        }
    }

    public void QueryProtectionStatus(out uint ProtectionState, out string ProtectedUntilTime)
    {
        _FhConfigMgr.QueryProtectionStatus(out ProtectionState, out ProtectedUntilTime);
    }

    protected bool _IsProtected(uint ProtectionState)
    {
        return ProtectionState == 0xFF;
    }

    public bool IsProtected
    {
        get
        {
            uint ProtectionState;
            string ProtectedUntilTime;
            QueryProtectionStatus(out ProtectionState, out ProtectedUntilTime);
            return _IsProtected(ProtectionState);
        }
    }
    public string ProtectionStatus
    {
        get
        {
            uint ProtectionState;
            string ProtectedUntilTime;
            QueryProtectionStatus(out ProtectionState, out ProtectedUntilTime);
            return ProtectedUntilTime;
        }
    }

    public void StartBackup()
    {
        FH_SERVICE_PIPE_HANDLE pipe = NativeConstants.INVALID_HANDLE_VALUE;

        try
        {
            // TODO: Safety-checks
            uint hr = NativeMethods.FhServiceOpenPipe(true, out pipe);
            if (hr != 0)
            {
                throw new COMException(string.Format("Error: FhServiceOpenPipe failed (hr=0x{0:%X})", hr));
            }
            hr = NativeMethods.FhServiceStartBackup(pipe, true);
            if (hr != 0)
            {
                throw new COMException(string.Format("Error: FhServiceStartBackup failed (hr=0x{0:%X})", hr));
            }
        }
        finally
        {
            if (pipe != NativeConstants.INVALID_HANDLE_VALUE) {
                NativeMethods.FhServiceClosePipe(pipe);
            }
        }
    }
  }

}
"@

function global:Start-FHBackup { 
 <#
    .Synopsis
        Manually starts a File History backup.
    .Description
        Triggers a File History backup (Windows 8/8.1/10) using public APIs.
        Assumes File History has been properly configured.
        Immediately returns control after invocation, you need to open the Control Panel to monitor progress.
        Throws an exception when an error occurs.
        See the C++ "File History Sample Setup Tool" example.
        Does not work in a non-native console (i.e., a 64-bit Windows OS running a 32-bit powershell console)
    .Example
        Start-FHBackup
#>

process {
   $mgr = New-Object Mchubby.Interop.FhConfigMgr
   $mgr.StartBackup()
}
}

