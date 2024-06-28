{ Copyright (C) 2024 by Bill Stewart (bstewart at iname.com)

  This program is free software: you can redistribute it and/or modify it under
  the terms of the GNU General Public License as published by the Free Software
  Foundation, either version 3 of the License, or (at your option) any later
  version.

  This program is distributed in the hope that it will be useful, but WITHOUT
  ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
  FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
  details.

  You should have received a copy of the GNU General Public License
  along with this program. If not, see https://www.gnu.org/licenses/.

}

unit WinShortcutRunas;

{$MODE OBJFPC}
{$MODESWITCH UNICODESTRINGS}

interface

uses
  windows;

type
  TFileChangeState = (
    NoChange = -1,
    Disabled = 0,
    Enabled = 1);

function GetShortcutRunas(const FileName: string; var Enabled: Boolean): DWORD;

function SetShortcutRunas(const FileName: string; const Enable, Force: Boolean;
  out FileChangeState: TFileChangeState): DWORD;

implementation

uses
  activex,
  comobj,
  shlobj;

const
  SID_IShellLinkDataList = '{45E2B4AE-B1C3-11D0-B92F-00A0C90312E1}';
  SLDF_RUNAS_USER = $00002000;

type
  IShellLinkDataList = interface
    [SID_IShellLinkDataList]
    function AddDataBlock(pDataBlock: Pointer): HRESULT; stdcall;
    function CopyDataBlock(dwSig: DWORD; var ppDataBlock: Pointer): HRESULT; stdcall;
    function RemoveDataBlock(dwSig: DWORD): HRESULT; stdcall;
    function GetFlags(out pdwFlags: DWORD): HRESULT; stdcall;
    function SetFlags(dwFlags: DWORD): HRESULT; stdcall;
  end;

type
  TShortcutFile = class
  strict private
  var
    Flags: DWORD;
    Error: HRESULT;
    IObject: IUnknown;
    IPFile: IPersistFile;
    ISDataList: IShellLinkDataList;
  public
    property Status: HRESULT read Error;
    function GetShortcutRunas(): Boolean;
    procedure SetShortcutRunas(const Enabled: Boolean);
    constructor Create(const FileName: string; const ReadOnly: Boolean);
    destructor Destroy(); override;
  end;

function TShortcutFile.GetShortcutRunas(): Boolean;
begin
  if Error = ERROR_SUCCESS then
    result := Flags and SLDF_RUNAS_USER <> 0
  else
    result := false;
end;

procedure TShortcutFile.SetShortcutRunas(const Enabled: Boolean);
var
  NewFlags: DWORD;
begin
  if Error <> ERROR_SUCCESS then
    exit;
  if Enabled then
    NewFlags := Flags or SLDF_RUNAS_USER
  else
    NewFlags := Flags and (not SLDF_RUNAS_USER);
  if NewFlags <> Flags then
  begin  // only call SetFlags if changed
    Error := ISDataList.SetFlags(NewFlags);
    if Error = ERROR_SUCCESS then
      Flags := NewFlags;
  end;
  if Error = ERROR_SUCCESS then
    Error := IPFile.Save(nil, true);
end;

constructor TShortcutFile.Create(const FileName: string; const ReadOnly: Boolean);
var
  OpenMode: DWORD;
begin
  Flags := 0;
  IObject := CreateComObject(CLSID_ShellLink);
  IPFile := IObject as IPersistFile;
  ISDataList := IObject as IShellLinkDataList;
  if ReadOnly then
    OpenMode := STGM_READ
  else
    OpenMode := STGM_READWRITE;
  Error := IPFile.Load(PChar(FileName), OpenMode);
  if Error = ERROR_SUCCESS then
    Error := ISDataList.GetFlags(Flags);
end;

destructor TShortcutFile.Destroy();
begin
end;

function GetWin32Error(const ErrorCode: HRESULT): DWORD;
begin
  if ErrorCode < 0 then
    result := DWORD(ErrorCode) and $0000FFFF
  else
    result := DWORD(ErrorCode);
end;

function GetShortcutRunas(const FileName: string; var Enabled: Boolean): DWORD;
var
  ShortcutFile: TShortcutFile;
begin
  ShortcutFile := TShortcutFile.Create(FileName, true);  // true = open as read-only
  result := GetWin32Error(ShortcutFile.Status);
  if result = ERROR_SUCCESS then
    Enabled := ShortcutFile.GetShortcutRunas();
  ShortcutFile.Destroy();
end;

function SetShortcutRunas(const FileName: string; const Enable, Force: Boolean;
  out FileChangeState: TFileChangeState): DWORD;
var
  Enabled: Boolean;
  ShortcutFile: TShortcutFile;
begin
  if not Force then
  begin
    result := GetShortcutRunas(FileName, Enabled);
    if result <> ERROR_SUCCESS then
      exit;
    if Enable = Enabled then
    begin
      FileChangeState := TFileChangeState.NoChange;
      exit;
    end;
  end;
  ShortcutFile := TShortcutFile.Create(FileName, false);  // false = open as read-write
  result := GetWin32Error(ShortcutFile.Status);
  if result = ERROR_SUCCESS then
  begin
    ShortcutFile.SetShortcutRunas(Enable);
    result := GetWin32Error(ShortcutFile.Status);
    if result = ERROR_SUCCESS then
    begin
      if Enable then
        FileChangeState := TFileChangeState.Enabled
      else
        FileChangeState := TFileChangeState.Disabled;
    end;
  end;
  ShortcutFile.Destroy();
end;

initialization

finalization

end.
