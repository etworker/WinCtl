// 最大化、最小化、恢复、关闭、层叠、平铺、隐藏、重现、只显指定窗口
// WinCtl.exe 参数1 参数2
// 参数1 = (Max | Min | Restore | Close | Move | Top | UnTop | Hide | UnHide |
//          Alpha | UnAlpha | HideTitle | UnHideTitle | Round | UnRound |
//          MinAll |UnMinAll | Cas | TH | TV | ShowOnly | GetHandle | GetCaption | GetClass)
// 参数2 = [ddd | Handle(=|!=|~|!~)ddd | Caption(=|!=|~|!~)"xxx" | Class(=|!=|~|!~)"xxx" | ALL]

// 参数1
// -----------
// Max (最大化)
// Min (最小化)
// Restore (恢复)
// Close (关闭窗体)
// Move (移动窗体)
// Top (窗体置顶)
// UnTop (窗体取消置顶)
// Hide (窗体隐藏)
// UnHide (窗体恢复)
// Alpha (窗体半透明，透明度)
// UnAlpha (窗体取消半透明)
// HideTitle (隐藏标题)
// UnHideTitle (恢复标题)
// Round (圆角窗体，圆角半径)
// UnRound (取消圆角窗体)
// MinAll (显示桌面)
// UnMinAll (恢复桌面)
// Cas (窗体层叠)
// TH (窗体横向平铺)
// TV (窗体纵向平铺)
// ShowOnly (仅显示指定窗体，其余最小化)
// GetHandle (取得当前窗体句柄，放入剪贴板)
// GetCaption (取得当前窗体标题，放入剪贴板)
// GetClass (取得当前窗体类名，放入剪贴板)



// 参数2
// -----------
// = 等于
// != 不等于
// ~ 正则表达式匹配
//     如 "记事本" 可以匹配 "无标题-记事本"
//     如 "a.*b" 可以匹配 "aqecbd", ".*" 代表任意个数的任意字符
// !~ 正则表达式不匹配
// ddd 表示数字
// xxx 表示字符

program WinCtl;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  untWinController in 'Unit\untWinController.pas',
  untOption in 'Unit\untOption.pas',
  untUtility in 'Unit\untUtility.pas',
  untClipboard in 'Unit\untClipboard.pas',
  untLogger in 'Unit\untLogger.pas';

{$R *.res}

var
  WinController: TWinController;
  i: Cardinal;

procedure ShowUsage;
begin
  Writeln('');
  Writeln('');
  Writeln('NAME');
  Writeln('  WinCtl - Contrl Window to ');
  Writeln('           Max/Min/Tile/Hide/UnHide/Top/UnTop/ShowOnly/');
  Writeln('           GetCaption/GetClass');
  Writeln('');
  Writeln('');
  Writeln('VERSION');
  Writeln('  V' + VERSION);
  Writeln('');
  Writeln('');
  Writeln('USAGE');
  Writeln('  Max Window:');
  Writeln('    WinCtl.exe Max ddd');
  Writeln('    WinCtl.exe Max Handle=ddd      // Handle of Window is ddd');
  Writeln('    WinCtl.exe Max Handle!=ddd     // Handle of Window is not ddd');
  Writeln('    WinCtl.exe Max Handle~ddd      // Handle of Window is ddd (Regex Mode)');
  Writeln('    WinCtl.exe Max Handle!~ddd     // Handle of Window is not ddd (Regex Mode)');
  Writeln('    WinCtl.exe Max Caption="xxx"   // Caption of Window is xxx');
  Writeln('    WinCtl.exe Max Caption!="xxx"  // Caption of Window is not xxx');
  Writeln('    WinCtl.exe Max Caption~"xxx"   // Caption of Window is xxx (Regex Mode)');
  Writeln('    WinCtl.exe Max Caption!~"xxx"  // Caption of Window is not xxx (Regex Mode)');
  Writeln('    WinCtl.exe Max Class="xxx"     // Class of Window is xxx');
  Writeln('    WinCtl.exe Max Class!="xxx"    // Class of Window is not xxx');
  Writeln('    WinCtl.exe Max Class~"xxx"     // Class of Window is xxx (Regex Mode)');
  Writeln('    WinCtl.exe Max Class!~"xxx"    // Class of Window is not xxx (Regex Mode)');
  Writeln('    WinCtl.exe Max ALL             // All Windows in Taskbar');
  Writeln('');
  Writeln('  Min Window:');
  Writeln('    (Similar to "Max")');
  Writeln('');
  Writeln('  Restore Window:');
  Writeln('    (Similar to "Max")');
  Writeln('');
  Writeln('  Show the Only Window:');
  Writeln('    (Similar to "Max")');
  Writeln('');
  Writeln('  Close Window:');
  Writeln('    (Similar to "Max")');
  Writeln('    WinCtl.exe Move 12321 [LEFT|RIGHT|UP|DOWN|left, top, width, height]');
  Writeln('');

  // winctl.exe Move 657202 Left             // 左半边
  // winctl.exe Move 657202 Right            // 右半边
  // winctl.exe Move 657202 Top              // 上半边
  // winctl.exe Move 657202 Bottom           // 下半边
  // winctl.exe Move 657202 Left+Top         // 左上 1/4
  // winctl.exe Move 657202 Left+Bottom      // 左下 1/4
  // winctl.exe Move 657202 Right+Top        // 右上 1/4
  // winctl.exe Move 657202 Right+Bottom     // 右下 1/4
  // winctl.exe Move 657202 100,200,300,400  // 左(像素)，上(像素)，宽(像素)，高(像素)
  // winctl.exe Move 657202 0.2,0.3,0.4,0.5  // 左(比例),上(比例)，宽(比例)，高(比例)
  // winctl.exe Move 657202 100,200,0.5,0.5  // 左(像素),上(像素)，宽(比例)，高(比例)
  Writeln('  Move Window:');
  Writeln('    (Similar to "Max")');
  Writeln('');
  Writeln('  Hide Window:');
  Writeln('    (Similar to "Max")');
  Writeln('');
  Writeln('  UnHide Window:');
  Writeln('    (Similar to "Max")');
  Writeln('    WinCtl.exe UnHide              // UnHide the latest hided window');
  Writeln('');
  Writeln('  Alpha Window:');
  Writeln('    (Similar to "Max"), and add one param: Alpha(From 0 to 255)');
  Writeln('    WinCtl.exe Alpha 12321 150     // Alpha = 150 for the specific window');
  Writeln('');
  Writeln('  UnAlpha Window:');
  Writeln('    (Similar to "Max")');
  Writeln('    WinCtl.exe UnAlpha             // UnAlpha the latest Alpha window');
  Writeln('');
//  Writeln('  RoundRect Window:');
//  Writeln('    (Similar to "Max"), and add one param: RoundBorderRadius(Greater than 0)');
//  Writeln('    WinCtl.exe Round 12321 15      // RoundBorderRadius = 15 for the specific window');
//  Writeln('');
//  Writeln('  UnRoundRect Window:');
//  Writeln('    (Similar to "Max")');
//  Writeln('    WinCtl.exe UnRound             // UnRound the latest Rounded window');
//  Writeln('');
  Writeln('  Top Window:');
  Writeln('    (Similar to "Max")');
  Writeln('');
  Writeln('  UnTop Window:');
  Writeln('    (Similar to "Max")');
  Writeln('    WinCtl.exe UnTop               // UnTop the latest hided window');
  Writeln('');
  Writeln('  Min All Window:');
  Writeln('    WinCtl.exe MinAll');
  Writeln('');
  Writeln('  Undo Min All Window:');
  Writeln('    WinCtl.exe UnMinAll');
  Writeln('');
  Writeln('  Cascade Windows:');
  Writeln('    WinCtl.exe Cas');
  Writeln('');
  Writeln('  Tile Horizontally Windows:');
  Writeln('    WinCtl.exe TH');
  Writeln('');
  Writeln('  Tile Vertically Windows:');
  Writeln('    WinCtl.exe TV');
  Writeln('');
  Writeln('  Get Caption of the Window:');
  Writeln('    WinCtl.exe GetCaption ddd');
  Writeln('');
  Writeln('  Get ClassName of the Window:');
  Writeln('    WinCtl.exe GetClass ddd');
  Writeln('');
  Writeln('');
  Writeln('DESCRIPTION');
  Writeln('  Regex Mode');
  Writeln('    Regular Expression Mode');
  Writeln('');
  Writeln('');
  Writeln('AUTHOR');
  Writeln('  ET Worker - JourneyBoy@GMail.com');
end;

begin
  InitLogger(False, False, False);

  WinController := TWinController.Create;

  WinController.GetForegroundWindowHandle;

  //WinController.Test; Exit;

  // 如果没有参数，则显示使用方法
  if ParamCount = 0 then
  begin
    ShowUsage;
    Exit;
  end;

  // 第1个参数是命令
  if ParamCount >= 1 then
    if not WinController.ReadWinCmdList(ParamStr(1)) then
    begin
      ShowUsage;
      Exit;
    end;

  // 第3个参数是其他参数
  // 因为读取第2个参数时，需要使用到第3个参数，所以先读第3个参数
  if ParamCount >= 3 then
    WinController.Param3 := ParamStr(3);

  // 第2个参数是窗体参数
  if ParamCount >= 2 then
    if Length(WinController.WinCmdList) > 0 then
      for i := 0 to High(WinController.WinCmdList) do        
      begin
        if not WinController.ReadWinParam(WinController.WinCmdList[i], ParamStr(2)) then
        begin
          ShowUsage;
          Exit;
        end;
      end;

  WinController.HandleCmdList;

  WinController.Free;
end.
