unit gwiopm;
{-----------------------------------------------
Functions for interacting with gwiopm I/O-permissions-map "driver",
in order to give cpu I/O instructions permission to operate
via the I/O permissions mechanism under Win NT.

Revisions
---------
98-05-23 GW original

Copyright Graham Wideman
------------------------
This module is distributed as freeware, and may be freely used for any purpose.
I would appreciate a credit notice if this is useful in your work. Thanks.

Note that this work was greatly aided by demo code from:
Dale Roberts      (giveio.sys)
Paula Tomlinson   (LOADDRV)

------------------------------------------------}
interface
uses windows, winsvc;

Const
  IOPM_SIZE = $2000;
Type
  TIOPM = array[0..IOPM_SIZE] of byte;
  PIOPM = ^TIOPM;

  TGWIOPM_Driver = class
  private
    HomeDir         : string;
    DriverDir       : string;
    DriverName      : string;
    DriverPath      : string; // the whole thing
    hSCMan          : SC_HANDLE; // Service Control Manager
    hDevice         : SC_HANDLE; // Handle for device

    //---------------------------------------
    // private multi-func funcs called by public ones
    //---------------------------------------
    function IOCTL_IOPMD_Misc1(var RetVal: DWORD; Cmd: integer):  DWORD;
    function IOCTL_IOPMD_GET_SET_LIOPM(Addr: Word; var B: byte; cmd: integer): DWORD;

  public
    LastOutBuf      : longint;
    constructor Create;

    //---------------------------------------
    // Interact with Service Control Manager
    //---------------------------------------
    function OpenSCM: DWORD;
    function CloseSCM: DWORD;

    //---------------------------------------
    // Install/Start/Stop/Remove driver
    //---------------------------------------
    function Install(newdriverpath: string): DWORD; { use '' for default }
    function Start:   DWORD;
    function Stop:    DWORD;
    function Remove:  DWORD;

    //--------------------------------
    // Device Open/Close
    //--------------------------------
    function DeviceOpen: DWORD;   // get a valid hDevice
    function DeviceClose: DWORD;

    //--------------------------------
    // IO Permission Map functions
    //--------------------------------
    // Test functions
    function IOCTL_IOPMD_READ_TEST(var RetVal: DWORD): DWORD;
    function IOCTL_IOPMD_READ_VERSION(var RetVal: DWORD): DWORD;
    // Manipulate driver's local IOPM (LIOPM)
    function IOCTL_IOPMD_CLEAR_LIOPM:  DWORD;    // set "local" map to block perm for all I/O addr
    function IOCTL_IOPMD_SET_LIOPM(Addr: Word; B: byte): DWORD;  // set a byte (8 ports-worth) in LIOPM
    function IOCTL_IOPMD_GET_LIOPMB(Addr: Word; var B: byte): DWORD;  // get a byte from LIOPM (diagnostic)
    function IOCTL_IOPMD_GET_LIOPMA(var A: TIOPM): DWORD;  // get entire LIOPM array (diagnostic)

    function LIOPM_Set_Ports(BeginPort: word; EndPort: word; Enable: Boolean): DWORD;

    // Interact with kernel IOPM (KIOPM)
    function IOCTL_IOPMD_ACTIVATE_KIOPM: DWORD;              // copy LIOPM to be active map
    function IOCTL_IOPMD_DEACTIVATE_KIOPM: DWORD;            // tell kernel to forget map
    function IOCTL_IOPMD_QUERY_KIOPM: DWORD;                  // query KIOPM to LIOPM
    //--------------------------------
    function ErrorLookup(ErrorNum: DWORD): string;
  end;

Const
  DEVICE_NAME_STRING	= 'gwiopm'; // in application's home directory

  // Device type           -- in the "User Defined" range."
  IOPMD_TYPE = $F100;

  // The IOCTL function codes from 0x800 to 0xFFF are for non-Microsoft use.
  // Test functions
  IOCMD_IOPMD_READ_TEST        = $900;
  IOCMD_IOPMD_READ_VERSION     = $901;
  // Manipulate driver's local IOPM (LIOPM)
  IOCMD_IOPMD_CLEAR_LIOPM      = $910;
  IOCMD_IOPMD_SET_LIOPM        = $911;
  IOCMD_IOPMD_GET_LIOPMB       = $912;
  IOCMD_IOPMD_GET_LIOPMA       = $913;
  // Interact with kernel IOPM (KIOPM)
  IOCMD_IOPMD_ACTIVATE_KIOPM   = $920;
  IOCMD_IOPMD_DEACTIVATE_KIOPM = $921;
  IOCMD_IOPMD_QUERY_KIOPM      = $922;

Var
  GWIOPM_Driver: TGWIOPM_Driver;

//-------------------------------------------
implementation
//-------------------------------------------
uses sysutils;

Const // from ntddk
// Service Types (Bit Mask)
  SERVICE_KERNEL_DRIVER        =  $00000001;
  SERVICE_FILE_SYSTEM_DRIVER   =  $00000002;
  SERVICE_ADAPTER              =  $00000004;
  SERVICE_RECOGNIZER_DRIVER    =  $00000008;

  SERVICE_DRIVER               =  SERVICE_KERNEL_DRIVER OR
                                  SERVICE_FILE_SYSTEM_DRIVER OR
                                  SERVICE_RECOGNIZER_DRIVER;

  SERVICE_WIN32_OWN_PROCESS    =  $00000010;
  SERVICE_WIN32_SHARE_PROCESS  =  $00000020;
  SERVICE_WIN32                =  SERVICE_WIN32_OWN_PROCESS OR
                                  SERVICE_WIN32_SHARE_PROCESS;

  SERVICE_INTERACTIVE_PROCESS  =  $00000100;

  SERVICE_TYPE_ALL             =  SERVICE_WIN32   OR
                                  SERVICE_ADAPTER OR
                                  SERVICE_DRIVER  OR
                                  SERVICE_INTERACTIVE_PROCESS;
// Start Type
  SERVICE_BOOT_START           =  $00000000;
  SERVICE_SYSTEM_START         =  $00000001;
  SERVICE_AUTO_START           =  $00000002;
  SERVICE_DEMAND_START         =  $00000003;
  SERVICE_DISABLED             =  $00000004;

// Error control type
  SERVICE_ERROR_IGNORE         =  $00000000;
  SERVICE_ERROR_NORMAL         =  $00000001;
  SERVICE_ERROR_SEVERE         =  $00000002;
  SERVICE_ERROR_CRITICAL       =  $00000003;

Type
  TErrorMsg = record
    Num: integer;
    Msg: string;
  end;

Const
  ErrorMsgCt = 30;
  ERROR_SCM_CANT_CONNECT = 9998;
  ERROR_NO_DEVICE_HANDLE = 9997;
  ERROR_GW_BUFFER_TOO_SMALL = 9997;
  ERROR_UNEXPECTED = 9999;

  ErrorMsgs: array[1..ErrorMsgCt] of TErrorMsg = (
    (Num: ERROR_SUCCESS                   ; Msg: 'Operation was successful'),
    (Num: ERROR_INVALID_FUNCTION          ; Msg: 'Invalid Function'),
    (Num: ERROR_ACCESS_DENIED             ; Msg: 'Access denied'),
    (Num: ERROR_CIRCULAR_DEPENDENCY       ; Msg: 'Circular dependency'),
    (Num: ERROR_DATABASE_DOES_NOT_EXIST   ; Msg: 'Database doesn''t exist'),
    (Num: ERROR_DEPENDENT_SERVICES_RUNNING; Msg: 'Dependent services running'),
    (Num: ERROR_DUP_NAME                  ; Msg: 'Display name already exists'),
    (Num: ERROR_INVALID_HANDLE            ; Msg: 'Invalid handle'),
    (Num: ERROR_INVALID_NAME              ; Msg: 'Invalid service name'),
    (Num: ERROR_INVALID_PARAMETER         ; Msg: 'Invalid Parameter'),
    (Num: ERROR_INVALID_SERVICE_ACCOUNT   ; Msg: 'User account doesn''t exist'),
    (Num: ERROR_INVALID_SERVICE_CONTROL   ; Msg: 'Invalid service control code'),
    (Num: ERROR_PATH_NOT_FOUND            ; Msg: 'Path not found'),
    (Num: ERROR_SERVICE_ALREADY_RUNNING   ; Msg: 'Service already running'),
    (Num: ERROR_SERVICE_CANNOT_ACCEPT_CTRL; Msg: 'Service can''t accept control'),
    (Num: ERROR_SERVICE_DATABASE_LOCKED   ; Msg: 'The database is locked'),
    (Num: ERROR_SERVICE_DEPENDENCY_DELETED; Msg: 'Depends on nonexistant service'),
    (Num: ERROR_SERVICE_DEPENDENCY_FAIL   ; Msg: 'Depends on service that failed'),
    (Num: ERROR_SERVICE_DISABLED          ; Msg: 'Service has been disabled'),
    (Num: ERROR_SERVICE_DOES_NOT_EXIST    ; Msg: 'Service doesn''t exist'),
    (Num: ERROR_SERVICE_EXISTS            ; Msg: 'Service already exists'),
    (Num: ERROR_SERVICE_LOGON_FAILED      ; Msg: 'Service couldn''t be logged on'),
    (Num: ERROR_SERVICE_MARKED_FOR_DELETE ; Msg: 'Service marked for deletion'),
    (Num: ERROR_SERVICE_NO_THREAD         ; Msg: 'Couldn''t create thread'),
    (Num: ERROR_SERVICE_NOT_ACTIVE        ; Msg: 'Service hasn''t been started'),
    (Num: ERROR_SERVICE_REQUEST_TIMEOUT   ; Msg: 'Service timed out'),
    (Num: ERROR_GW_BUFFER_TOO_SMALL       ; Msg: 'Buffer too small'),
    (Num: ERROR_NO_DEVICE_HANDLE          ; Msg: 'No device handle'),
    (Num: ERROR_SCM_CANT_CONNECT          ; Msg: 'Can''t connect to Service Control Manager'),
    (Num: ERROR_UNEXPECTED                ; Msg: 'An unexpected error occured')
  );

//-----------------------------------------
function TGWIOPM_Driver.ErrorLookup(ErrorNum: DWORD): string;
//-----------------------------------------
Var
  S: string;
  N: integer;
label foundit;
Begin
  If Error <> ERROR_SUCCESS then
    S := 'Error: ' + IntToStr(ErrorNum) + ': ';
    
  For N := 1 to ErrorMsgCt do
  Begin
    if ErrorNum = ErrorMsgs[N].Num then
    Begin
      goto foundit;
    end;
  end;
foundit:
  If N > ErrorMsgCt then N := ErrorMsgCt;
  S := S + ErrorMsgs[N].Msg;
  result := S;
end;

//----------------------------------------------------------
// IOCTL codes
//----------------------------------------------------------
function CTL_CODE(DeviceType: integer; func: integer; meth: integer; access: integer): DWORD;
Begin
  result := (DeviceType shl 16) or (Access shl 14) or (func shl 2) or (meth);
end;

Const
  // Buffering method for user-mode app talking to drive
  METHOD_BUFFERED    = 0;
  METHOD_IN_DIRECT   = 1;
  METHOD_OUT_DIRECT  = 2;
  METHOD_NEITHER     = 3;

  // Define the access allowed
  FILE_ANY_ACCESS    = 0;
  FILE_READ_ACCESS   = 1;     // file & pipe
  FILE_WRITE_ACCESS  = 2;     // file & pipe


//-----------------------------------------
constructor TGWIOPM_Driver.Create;
//-----------------------------------------
Begin
  hSCMan  := 0;
  hDevice := INVALID_HANDLE_VALUE;
  HomeDir := ExtractFilePath(ParamStr(0));
  DriverName  := DEVICE_NAME_STRING;
    // default driver name needed by stop/remove if install wasn't executed
    // this run (ie: driver already installed
end;

//-------------------------------------------
function TGWIOPM_Driver.OpenSCM:  DWORD;
//-------------------------------------------
Begin
  result := ERROR_SUCCESS;
  hSCMan := OpenSCManager(nil, nil, SC_MANAGER_ALL_ACCESS);
  if hSCMan = 0 then result := ERROR_SCM_CANT_CONNECT;
end;

//-------------------------------------------
function TGWIOPM_Driver.CloseSCM:  DWORD;
//-------------------------------------------
Begin
  result := ERROR_SUCCESS;
  CloseServiceHandle(hSCMan);
  hSCMan := 0;
end;

//-----------------------------------------
function TGWIOPM_Driver.Install(newdriverpath: string): DWORD; { use '' for default }
//-----------------------------------------
Var
  hService: SC_HANDLE;
  dwStatus: DWORD;
Begin
  hService := 0;
  dwStatus := 0;

  If newdriverpath = '' then
  Begin
    DriverDir   := HomeDir;
    DriverName  := DEVICE_NAME_STRING;
  end else
  Begin
    DriverDir  := ExtractFilePath(driverpath);
    DriverName := ExtractFileName(driverpath);
  end;
  DriverPath  := DriverDir + DriverName + '.sys';

   // add to service control manager's database
   hService := CreateService(hSCMan, PChar(DriverName),PChar(DriverName),
              SERVICE_ALL_ACCESS, SERVICE_KERNEL_DRIVER, SERVICE_DEMAND_START,
              SERVICE_ERROR_NORMAL, PChar(DriverPath),
              nil, nil, nil, nil, nil);
   if (hService = 0) then
   Begin
     dwStatus := GetLastError();
   end else
   Begin
     CloseServiceHandle(hService);
   end;

   result := dwStatus;
end;

//-------------------------------------------
function TGWIOPM_Driver.Start:   DWORD;
//-------------------------------------------
Var
  hService: SC_HANDLE;
  dwStatus: DWORD;
  lpServiceArgVectors: PChar;
  temp: LongBool;
Begin
  hService := 0;
  dwStatus := 0;
  lpServiceArgVectors := nil;

   // get a handle to the service
   hService := OpenService(hSCMan, PChar(DriverName), SERVICE_ALL_ACCESS);
   if hService <> 0 then
   Begin
      // start the driver
      temp := StartService(hService, 0, PChar(lpServiceArgVectors));
      if not temp then dwStatus := GetLastError();
   end else dwStatus := GetLastError();

   if (hService <> 0) then CloseServiceHandle(hService);
   result := dwStatus;
end;

//-------------------------------------------
function TGWIOPM_Driver.Stop:    DWORD;
//-------------------------------------------
Var
  hService: SC_HANDLE;
  dwStatus: DWORD;
  serviceStatus: TServiceStatus;
  temp: LongBool;
Begin
  hService := 0;
  dwStatus := 0;

  // get a handle to the service
  hService := OpenService(hSCMan, PChar(DriverName), SERVICE_ALL_ACCESS);
  if hService <> 0 then
  Begin
     // stop the driver
     temp := ControlService(hService, SERVICE_CONTROL_STOP, serviceStatus);
     if not temp then dwStatus := GetLastError();
  end else dwStatus := GetLastError();

  if (hService <> 0) then CloseServiceHandle(hService);
  result := dwStatus;
end;

//-------------------------------------------
function TGWIOPM_Driver.Remove:  DWORD;
//-------------------------------------------
Var
  hService: SC_HANDLE;
  dwStatus: DWORD;
  temp: LongBool;
Begin
  hService := 0;
  dwStatus := 0;

  dwStatus := Stop;  // ignore result
  dwStatus := 0;

  // get a handle to the service
  hService := OpenService(hSCMan, PChar(DriverName), SERVICE_ALL_ACCESS);
  if hService <> 0 then
  Begin
     temp := DeleteService(hService);
     if not temp then dwStatus := GetLastError();
  end else dwStatus := GetLastError();

  if (hService <> 0) then CloseServiceHandle(hService);
  result := dwStatus;
end;

//=============================================================
// Device Open/Close functions
//=============================================================

//-------------------------------------------
function TGWIOPM_Driver.DeviceOpen:  DWORD;
//-------------------------------------------
Var
  dwStatus: DWORD;
Begin
  dwStatus := 0;

  if hDevice <> INVALID_HANDLE_VALUE then DeviceClose;

  // get a handle to the device
  hDevice := CreateFile(
             { lpFileName: PChar            } '\\.\'+ DEVICE_NAME_STRING,
             { dwDesiredAccess: integer     } GENERIC_READ or GENERIC_WRITE,
             { dwShareMode: Integer         } 0,
             { lpSecurityAttributes         } PSECURITY_DESCRIPTOR(nil),
             { dwCreationDisposition: DWORD } OPEN_EXISTING,
             { dwFlagsAndAttributes: DWORD  } FILE_ATTRIBUTE_NORMAL,
             { hTemplateFile: THandle       } 0);

  if hDevice = INVALID_HANDLE_VALUE then
  Begin
    dwStatus := GetLastError();
  end;

  result := dwStatus;
end;

//-------------------------------------------
function TGWIOPM_Driver.DeviceClose:  DWORD;
//-------------------------------------------
Var
  dwStatus: DWORD;
Begin
  dwStatus := 0;
  if (hDevice <> INVALID_HANDLE_VALUE) then CloseHandle(hDevice);
  hDevice := INVALID_HANDLE_VALUE;
  result := dwStatus; { assume that it went OK? }
end;

//=================================================================
// IO Permission Map functions
//=================================================================
Const
  GWIO_PARAMCOUNT     = 3;
  GWIO_BYTES_IN  = GWIO_PARAMCOUNT * 4;
  GWIO_BYTES_OUT = GWIO_PARAMCOUNT * 4;

Type
  TGWIO_PARAMS = array[0..GWIO_PARAMCOUNT-1] of longint;

//-------------------------------------------
function TGWIOPM_Driver.IOCTL_IOPMD_Misc1(var RetVal: DWORD; Cmd: integer):  DWORD;
//-------------------------------------------
Var
  dwStatus: DWORD;
  temp: LongBool;
  BytesReturned: DWORD;
  MyControlCode: DWORD;
  InBuf:  TGWIO_PARAMS;
  OutBuf: TGWIO_PARAMS;

Begin
  dwStatus := 0;
  RetVal := 0;
  InBuf[0] := 0;
  InBuf[1] := 0;
  InBuf[2] := 0;

  if hDevice = INVALID_HANDLE_VALUE then
  Begin
    dwStatus := ERROR_NO_DEVICE_HANDLE;
  end else
  Begin
    MyControlCode := CTL_CODE(IOPMD_TYPE, Cmd , METHOD_BUFFERED, FILE_ANY_ACCESS);

    BytesReturned := 0;
    temp := DeviceIoControl(hDevice, MyControlCode ,
            { in buffer  (to driver)   }  @InBuf,  GWIO_BYTES_IN,
            { out buffer (from driver) }  @OutBuf, GWIO_BYTES_OUT,
            BytesReturned, nil);
    if temp then
    Begin
      RetVal := OutBuf[0];
    end else
    Begin
      dwStatus := GetLastError();
    end;
  end;

  result := dwStatus;
end;

//-------------------------------------------
function TGWIOPM_Driver.IOCTL_IOPMD_READ_TEST(var RetVal: DWORD):  DWORD;
//-------------------------------------------
Begin
  result := IOCTL_IOPMD_Misc1(RetVal, IOCMD_IOPMD_READ_TEST);
end;

//-------------------------------------------
function TGWIOPM_Driver.IOCTL_IOPMD_READ_VERSION(var RetVal: DWORD):  DWORD;
//-------------------------------------------
Begin
  result := IOCTL_IOPMD_Misc1(RetVal, IOCMD_IOPMD_READ_VERSION);
end;

//-------------------------------------------
function TGWIOPM_Driver.IOCTL_IOPMD_CLEAR_LIOPM:  DWORD;
//-------------------------------------------
Var
  RetVal: DWORD;
Begin
  result := IOCTL_IOPMD_Misc1(RetVal, IOCMD_IOPMD_CLEAR_LIOPM );
end;

//-------------------------------------------
function TGWIOPM_Driver.IOCTL_IOPMD_GET_SET_LIOPM(Addr: Word; var B: byte; cmd: integer): DWORD;
//-------------------------------------------
Var
  dwStatus: DWORD;
  temp: LongBool;
  BytesReturned: DWORD;
  MyControlCode: DWORD;
  InBuf:  TGWIO_PARAMS;
  OutBuf: TGWIO_PARAMS;

Begin
  dwStatus := 0;

  if hDevice = INVALID_HANDLE_VALUE then
  Begin
    dwStatus := ERROR_NO_DEVICE_HANDLE;
  end else
  Begin
    MyControlCode := CTL_CODE(IOPMD_TYPE, cmd, METHOD_BUFFERED, FILE_ANY_ACCESS);

    InBuf[0] := Addr;
    InBuf[1] := B;

    BytesReturned := 0;
    temp := DeviceIoControl(hDevice, MyControlCode ,
            { in buffer  (to driver)   }  @InBuf,  GWIO_BYTES_IN,
            { out buffer (from driver) }  @OutBuf, GWIO_BYTES_OUT,
            BytesReturned, nil);
    if temp then
    Begin
      B := Lo(OutBuf[1]);
    end else dwStatus := GetLastError();
  end;

  result := dwStatus;
end;

//-------------------------------------------
function TGWIOPM_Driver.IOCTL_IOPMD_SET_LIOPM(Addr: Word; B: byte):  DWORD;
//-------------------------------------------
Begin
  result := IOCTL_IOPMD_GET_SET_LIOPM(Addr, B, IOCMD_IOPMD_SET_LIOPM);
end;

//-------------------------------------------
function TGWIOPM_Driver.IOCTL_IOPMD_GET_LIOPMB(Addr: Word; var B: byte):  DWORD;
//-------------------------------------------
Begin
  result := IOCTL_IOPMD_GET_SET_LIOPM(Addr, B, IOCMD_IOPMD_GET_LIOPMB);
end;

//-------------------------------------------
function TGWIOPM_Driver.IOCTL_IOPMD_GET_LIOPMA(var A: TIOPM): DWORD;
//-------------------------------------------
// get entire LIOPM array (diagnostic)
Var
  dwStatus: DWORD;
  temp: LongBool;
  BytesReturned: DWORD;
  MyControlCode: DWORD;
  InBuf:  TGWIO_PARAMS;
//  OutBuf: TGWIO_PARAMS;

Begin
  dwStatus := 0;

  if hDevice = INVALID_HANDLE_VALUE then
  Begin
    dwStatus := ERROR_NO_DEVICE_HANDLE;
  end else
  Begin
    MyControlCode := CTL_CODE(IOPMD_TYPE, IOCMD_IOPMD_GET_LIOPMA, METHOD_BUFFERED, FILE_ANY_ACCESS);

    InBuf[0] := 0;
    InBuf[1] := 0;
    InBuf[2] := 0;

    BytesReturned := 0;
    temp := DeviceIoControl(hDevice, MyControlCode ,
            { in buffer  (to driver)   }  @InBuf,  GWIO_BYTES_IN,
            { out buffer (from driver) }  @A,      IOPM_SIZE,
            BytesReturned, nil);
    if temp then
    Begin
      // do nothing
    end else dwStatus := GetLastError();
  end;
  result := dwStatus;
end;

//-------------------------------------------
function TGWIOPM_Driver.IOCTL_IOPMD_ACTIVATE_KIOPM:  DWORD;
//-------------------------------------------
Var
  RetVal: DWORD;
Begin
  result := IOCTL_IOPMD_Misc1(RetVal, IOCMD_IOPMD_ACTIVATE_KIOPM );
end;

//-------------------------------------------
function TGWIOPM_Driver.IOCTL_IOPMD_DEACTIVATE_KIOPM:  DWORD;
//-------------------------------------------
Var
  RetVal: DWORD;
Begin
  result := IOCTL_IOPMD_Misc1(RetVal, IOCMD_IOPMD_DEACTIVATE_KIOPM );
end;

//-------------------------------------------
function TGWIOPM_Driver.IOCTL_IOPMD_QUERY_KIOPM:  DWORD;
//-------------------------------------------
Var
  RetVal: DWORD;
Begin
  result := IOCTL_IOPMD_Misc1(RetVal, IOCMD_IOPMD_QUERY_KIOPM);
end;

//-------------------------------------------
function TGWIOPM_Driver.LIOPM_Set_Ports(BeginPort: word; EndPort: word; Enable: Boolean): DWORD;
//-------------------------------------------
Var
  PortNum: word;
  IOPM_Ix, IOPM_BitNum : integer;
  IOPM_Byte, Mask_Byte: byte;
  DriverResult: DWORD;
Label the_end;
Begin
  DriverResult := ERROR_SUCCESS;
  IOPM_Byte    := $FF;

  For PortNum := BeginPort to EndPort do
  Begin
    IOPM_Ix      := PortNum shr 3;  // 8 bits per byte;
    IOPM_BitNum  := PortNum and 7;  // lowest 3 bits;
    Mask_Byte     := 1 shl IOPM_BitNum;

    If (PortNum = BeginPort) or (IOPM_BitNum = 0) then
    Begin  // get IOPM byte
      DriverResult := IOCTL_IOPMD_GET_LIOPMB(IOPM_Ix, IOPM_Byte);
      if DriverResult <> ERROR_SUCCESS then goto the_end;
    end;

    If Enable then
    Begin   // set the bit to 0
      IOPM_Byte := IOPM_Byte and ($FF xor Mask_Byte);
    end else
    Begin   // set the bit to 1
      IOPM_Byte := IOPM_Byte or Mask_Byte;
    end;

    If (PortNum = EndPort) or (IOPM_BitNum = 7) then
    Begin    // Write out IOPM_Byte
      DriverResult := IOCTL_IOPMD_SET_LIOPM(IOPM_Ix, IOPM_Byte);
      if DriverResult <> ERROR_SUCCESS then goto the_end;
    end;
  end;
the_end:
  Result := DriverResult;
end;

//-------------------------------------------
initialization
//-------------------------------------------
  GWIOPM_Driver := TGWIOPM_Driver.Create;

//-------------------------------------------
finalization
//-------------------------------------------

end.

