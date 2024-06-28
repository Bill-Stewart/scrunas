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

program scrunas;

{$MODE OBJFPC}
{$MODESWITCH UNICODESTRINGS}
{$R *.res}

uses
  windows,
  wargcv,
  wgetopts,
  WindowsMessages,
  WinShortcutRunas;

const
  PROGRAM_NAME = 'scrunas';
  PROGRAM_COPYRIGHT = 'Copyright (C) 2024 by Bill Stewart';

type
  TActionParamGroup = (
    ActionParamList,
    ActionParamEnable,
    ActionParamDisable,
    ActionParamHelp);
  TActionParamSet = set of TActionParamGroup;

  TCommandLine = object
    ActionParamSet: TActionParamSet;
    Error: DWORD;
    Quiet: Boolean;
    Force: Boolean;
    FileName: string;
    procedure Parse();
  end;

function BoolToStr(const B: Boolean): string;
begin
  if B then
    result := 'enabled'
  else
    result := 'disabled';
end;

function IntToStr(const I: Integer): string;
begin
  Str(I, result);
end;

function GetFileVersion(const FileName: string): string;
var
  VerInfoSize, Handle: DWORD;
  pBuffer: Pointer;
  pFileInfo: ^VS_FIXEDFILEINFO;
  Len: UINT;
begin
  result := '';
  VerInfoSize := GetFileVersionInfoSizeW(PChar(FileName),  // LPCWSTR lptstrFilename
    Handle);                                               // LPDWORD lpdwHandle
  if VerInfoSize > 0 then
  begin
    GetMem(pBuffer, VerInfoSize);
    if GetFileVersionInfoW(PChar(FileName),  // LPCWSTR lptstrFilename
      Handle,                                // DWORD   dwHandle
      VerInfoSize,                           // DWORD   dwLen
      pBuffer) then                          // LPVOID  lpData
    begin
      if VerQueryValueW(pBuffer,  // LPCVOID pBlock
        '\',                      // LPCWSTR lpSubBlock
        pFileInfo,                // LPVOID  *lplpBuffer
        Len) then                 // PUINT   puLen
      begin
        with pFileInfo^ do
        begin
          result := IntToStr(HiWord(dwFileVersionMS)) + '.' +
            IntToStr(LoWord(dwFileVersionMS)) + '.' +
            IntToStr(HiWord(dwFileVersionLS));
        end;
      end;
    end;
    FreeMem(pBuffer);
  end;
end;

procedure Usage();
begin
  WriteLn(PROGRAM_NAME, ' ', GetFileVersion(ParamStr(0)), ' - ', PROGRAM_COPYRIGHT);
  WriteLn('This is free software and comes with ABSOLUTELY NO WARRANTY.');
  WriteLn();
  WriteLn('SYNOPSIS');
  WriteLn();
  WriteLn('Enables, disables, or displays the status of the ''Run as administrator'' setting');
  WriteLn('for a Windows shortcut (.lnk) file.');
  WriteLn();
  WriteLn('USAGE');
  WriteLn();
  WriteLn(PROGRAM_NAME, ' [[--enable | --disable] [--force]] [--quiet] <filename>');
  WriteLn();
  WriteLn('PARAMETERS');
  WriteLn();
  WriteLn('Specify --enable (or -e) to enable the ''Run as administrator'' setting, or');
  WriteLn('specify --disable (or -d) to disable the setting. The --force (or -f) option');
  WriteLn('updates the shortcut (.lnk) file even if the requested setting is already');
  WriteLn('configured.');
  WriteLn();
  WriteLn('Omit --enable (-e) or --disable (-d) to display the current state of the ''Run');
  WriteLn('as administrator'' option for the shortcut (.lnk) file.');
  WriteLn();
  WriteLn('The --quiet (or -q) option prevents output.');
  WriteLn();
  WriteLn('EXIT CODES');
  WriteLn();
  WriteLn('If you omit --enable (-e) or --disable (-d), the program will exit with an');
  WriteLn('exit code of 1 if the ''Run as administrator'' setting is enabled for the');
  WriteLn('shortcut file, or 0 otherwise. Any other exit code indicates an error.');
  WriteLn();
  WriteLn('If you specify either --enable (-e) or --disable (-d), the program will exit');
  WriteLn('with an exit code of 0 for success, or non-zero for an error.');
end;

procedure TCommandLine.Parse();
var
  Opts: array[1..6] of TOption;
  Opt: Char;
  I: Integer;
begin
  with Opts[1] do
  begin
    Name := 'disable';
    Has_arg := No_Argument;
    Flag := nil;
    Value := 'd';
  end;
  with Opts[2] do
  begin
    Name := 'enable';
    Has_arg := No_Argument;
    Flag := nil;
    Value := 'e';
  end;
  with Opts[3] do
  begin
    Name := 'force';
    Has_arg := No_Argument;
    Flag := nil;
    Value := 'f';
  end;
  with Opts[4] do
  begin
    Name := 'help';
    Has_arg := No_Argument;
    Flag := nil;
    value := 'h';
  end;
  with Opts[5] do
  begin
    Name := 'quiet';
    Has_arg := No_Argument;
    Flag := nil;
    value := 'q';
  end;
  with Opts[6] do
  begin
    Name := '';
    Has_arg := No_Argument;
    Flag := nil;
    Value := #0;
  end;
  ActionParamSet := [ActionParamList];
  Error := ERROR_SUCCESS;
  Quiet := false;
  OptErr := false;
  repeat
    Opt := GetLongOpts('defhq', @Opts[1], I);
    case Opt of
      'd':
      begin
        Include(ActionParamSet, ActionParamDisable);
        Exclude(ActionParamSet, ActionParamList);
      end;
      'e':
      begin
        Include(ActionParamSet, ActionParamEnable);
        Exclude(ActionParamSet, ActionParamList);
      end;
      'f':
      begin
        Force := true;
      end;
      'h':
      begin
        Include(ActionParamSet, ActionParamHelp);
        Exclude(ActionParamSet, ActionParamList);
      end;
      'q':
      begin
        Quiet := true;
      end;
    end;
  until Opt = EndOfOptions;
  FileName := ParamStr(OptInd);
  if (FileName = '') or (PopCnt(DWORD(ActionParamSet)) > 1) then
    ActionParamSet := [ActionParamHelp];
end;

var
  CommandLine: TCommandLine;
  RC: DWORD;
  Error, Enabled: Boolean;
  FileChangeState: TFileChangeState;
  OutStr: string;

begin
  CommandLine.Parse();
  if (ParamStr(1) = '/?') or (ActionParamHelp in CommandLine.ActionParamSet) then
  begin
    Usage();
    exit;
  end;

  if ActionParamList in CommandLine.ActionParamSet then
  begin
    RC := GetShortcutRunas(CommandLine.FileName, Enabled);
    Error := RC <> ERROR_SUCCESS;
    if not Error then
    begin
      if Enabled then
        RC := 1;
      OutStr := 'The ''Run as administrator'' shortcut setting is ' +
        BoolToStr(Enabled) + ' for shorcut file ''' + CommandLine.FileName + '''.';
    end
    else
    begin
      OutStr := 'Error opening shortcut file ''' + CommandLine.FileName +
        ''': ' + GetWindowsMessage(RC, true);
    end;
  end
  else
  begin
    if ActionParamDisable in CommandLine.ActionParamSet then
      Enabled := false
    else if ActionParamEnable in CommandLine.ActionParamSet then
      Enabled := true;
    RC := SetShortcutRunas(CommandLine.FileName, Enabled, Commandline.Force,
      FileChangeState);
    Error := RC <> ERROR_SUCCESS;
    if not Error then
    begin
      if FileChangeState = NoChange then
        OutStr := 'The ''Run as administrator'' setting for shortcut file ''' +
          CommandLine.FileName + ' is already ' + BoolToStr(Enabled) + '.'
      else
        OutStr := 'Successfully ' + BoolToStr(Enabled) +
          ' the ''Run as administrator'' setting for shortcut file ''' +
          CommandLine.FileName + '''.';
    end
    else
      OutStr := 'Error updating the ''Run as administrator'' setting for shortcut file '''
        + CommandLine.FileName + ''': ' + GetWindowsMessage(RC, true);
  end;

  if not CommandLine.Quiet then
    WriteLn(OutStr);

  ExitCode := Integer(RC);
end.
