Unit untWinController;

Interface

Uses
  Windows,
  SysUtils,
  Classes,
  ActiveX,
  Dialogs,
  ComObj,
  RegExpr,
  untLogger,
  untClipboard;

Type
  TWinCmdType     = (
    wcUnkown,
    wcMax,
    wcMin,
    wcRestore,
    wcMove,
    wcClose,
    wcTop,
    wcUnTop,
    wcHide,
    wcUnHide,
    wcAlpha,
    wcUnAlpha,
    wcHideTitle,
    wcUnHideTitle,
    wcRound,
    wcUnRound,
    wcCascade,
    wcTileHorizontally,
    wcTileVertically,
    wcMinAll,
    wcUnMinAll,
    wcShowOnly,
    wcGetHandle,
    wcGetCaption,
    wcGetClass,
    wcTest
    );
  TWinCmdTypeList = Array Of TWinCmdType;

  TShellFunction = (
    sf_MinimizeAll,       // 全部窗口最小化
    sf_UndoMinimizeALL,   // 窗口状态复原
    sf_CascadeWindows,    // 层叠窗口
    sf_TileHorizontally,  // 水平平铺窗口
    sf_TileVertically     // 垂直平铺窗口
    );

  TMatchRule = (
    mrEqual,
    mrNotEqual,
    mrRegexEqual,
    mrNotRegexEqual);

  TMatchType = (
    mtEmpty,
    mtSingle,
    mtAll,
    mtHandle,
    mtCaption,
    mtClass,
    mtNext
    );

  TWinParam = Record
    ParamName:  String;
    ParamValue: String;
    MatchType:  TMatchType;
    MatchRule:  TMatchRule;
  End;

  TWinProperty = Record
    Handle:    THandle;
    Caption:   String;
    ClassName: String;
  End;
  TWinPropertyList = Array Of TWinProperty;
  PWinPropertyList = ^TWinPropertyList;

Type
  TWinController = Class

  Public
    Constructor Create;
    Destructor Destroy; Override;

    Procedure Test;

    Function ReadWinCmdList(str: String): Boolean;
    Function ReadWinParam(WinCmd: TWinCmdType; str: String): Boolean;
    Function HandleCmd(WinCmd: TWinCmdType): Boolean;
    Function HandleCmdList: Boolean;

  Private
    //m_WinCmd: TWinCmdType;
    m_WinCmdList:      TWinCmdTypeList;
    m_WinList:         TStringList;
    m_HandleFileName:  String;
    m_HandleList:      TStringList;
    m_WinHandle:       Integer;
    m_WinParam:        TWinParam;
    m_Regex:           TRegExpr;
    m_WinPropertyList: TWinPropertyList;
    m_Param3:          String;
    m_ForegroundWindowHandle: Integer;

    Function ControlWindow(Handle: Integer; WinCmd: TWinCmdType; Param3: String = ''): Boolean;
    Function ShellFunction(ShellFunction: TShellFunction): Boolean;

    Procedure CompitableWithOldVersion(strList: TStringList);
    Function ParseWinParam(str: String; Var WinParam: TWinParam): Boolean;
    Function GetParamValue(str, Name: String): String;

    Function IsMatch(strText, strPattern: String; MatchRule: TMatchRule): Boolean;
    Procedure RefreshWinPropertyList;
    Function GetWinList(str: String; MatchType: TMatchType; MatchRule: TMatchRule): Boolean;
  Public
    //property WinCmd: TWinCmdType read m_WinCmd write m_WinCmd;
    Property WinCmdList: TWinCmdTypeList Read m_WinCmdList Write m_WinCmdList;
    Property Param3: String Read m_Param3 Write m_Param3;
    Property ActiveWindowHandle: Integer Read m_ForegroundWindowHandle Write m_ForegroundWindowHandle;

    Procedure GetForegroundWindowHandle;

  End;

Implementation

Uses
  untUtility;

Const
  WM_SYSCOMMAND = $112;
  WM_CLOSE  = $10;
  CLIP_FLAG = '{%clip}';
  HANDLE_FILE = 'HandleList.txt';
  HIDE_FLAG = 'Hide=';
  TOP_FLAG  = 'Top=';
  ALPHA_FLAG = 'Alpha=';
  TITLE_FLAG = 'Title=';
  ROUND_FLAG = 'Round=';

// 获得当前ConsoleWindow的句柄
Function GetConsoleWindow: HWND; Stdcall; External kernel32 Name 'GetConsoleWindow';

Function EnumWindowsInTaskBarFunc(Handle: THandle; pList: PWinPropertyList): Boolean; Stdcall;
Var
  Caption: Array[0..256] Of Char;
  ClassName: Array[0..256] Of Char;
Begin
  // 得到任务栏所有的窗口列表
  If (IsWindowVisible(Handle) {or IsIconic(Handle)}) And ((GetWindowLong(Handle, GWL_HWNDPARENT) = 0) Or
    (GetWindowLong(Handle, GWL_HWNDPARENT) = GetDesktopWindow)) And
    (GetWindowLong(Handle, GWL_EXSTYLE) And WS_EX_TOOLWINDOW = 0) Then
  Begin
    If (Handle <> GetConsoleWindow) And (GetWindowText(Handle, Caption, SizeOf(Caption) - 1) <> 0) And
      (GetClassName(Handle, ClassName, SizeOf(ClassName) - 1) <> 0) Then
    Begin
      SetLength(pList^, Length(pList^) + 1);

      pList^[High(pList^)].Handle  := Handle;
      pList^[High(pList^)].Caption := Caption;
      pList^[High(pList^)].ClassName := ClassName;
    End;
  End;

  Result := True;
End;

{ TWinController }

Procedure TWinController.CompitableWithOldVersion(strList: TStringList);
Var
  i: Cardinal;
Begin
  If strList.Count = 0 Then
    Exit;

  For i := 0 To strList.Count - 1 Do
  Begin
    strList.Strings[i] := Trim(strList.Strings[i]);

    // 如果某行是纯数字，说明是老版本产生的，将其前面加上 HIDE_FLAG
    If StrToIntDef(strList.Strings[i], -1) > 0 Then
      strList.Strings[i] := HIDE_FLAG + strList.Strings[i];
  End;
End;

{*------------------------------------------------------------------------------
  这是我的测试代码

  @param Handle 窗体句柄
  @param WinCmd 控制命令
  @param Param3 第3参数
  @see   WriteMessage
-------------------------------------------------------------------------------}

Function TWinController.ControlWindow(Handle: Integer; WinCmd: TWinCmdType; Param3: String = ''): Boolean;
Var
  Rect, ScreenRect: TRect;
  ScreenWidth, ScreenHeight: Integer;
  w, h: Integer;
  Index: Integer;
  List: TStringList;
  t: Integer;
  d: Array[1..4] Of Integer;
  i: Integer;
Begin
  Result := False;

  If Handle = 0 Then
    Exit;

  //TraceMsg('Handle = %d, Cmd = %d', [Handle, Ord(WinCmd)]);

  Case WinCmd Of
    wcMax:
    Begin
      If m_Param3 <> '' Then
      Begin
        //SendMessage(Handle, WM_SYSCOMMAND, SC_MAXIMIZE, 0);

        // 取得桌面大小
        SystemParametersInfo(SPI_GETWORKAREA, 0, @Rect, 0);

        Try
          List := TStringList.Create;
          SplitString(m_Param3, ',', List);

          // 格式为 “a,b,c,d”，允许丢三落四
          If List.Count >= 1 Then
            Rect.Left := Rect.Left - StrToIntDef(List.Strings[0], 0);

          If List.Count >= 2 Then
            Rect.Top := Rect.Top - StrToIntDef(List.Strings[1], 0);

          If List.Count >= 3 Then
            Rect.Right := Rect.Right + StrToIntDef(List.Strings[2], 0);

          If List.Count >= 4 Then
            Rect.Bottom := Rect.Bottom + StrToIntDef(List.Strings[3], 0);
        Finally
          List.Free;
        End;

        MoveWindow(Handle, Rect.Left, Rect.Top, Rect.Right - Rect.Left, Rect.Bottom - Rect.Top, True);
        //MoveWindow(Handle, -30, -30, 1300, 1300, True);
        //SetWindowPos(Handle,HWND_TOP,-130, -130, 2000, 2000,$40);
      End
      Else
        PostMessage(Handle, WM_SYSCOMMAND, SC_MAXIMIZE, 0);
    End;

    wcMin:
      PostMessage(Handle, WM_SYSCOMMAND, SC_MINIMIZE, 0);

    wcRestore:
      PostMessage(Handle, WM_SYSCOMMAND, SC_RESTORE, 0);

    wcClose:
      PostMessage(Handle, WM_CLOSE, 0, 0);

    wcMove:
    Begin
      If m_Param3 <> '' Then
      Begin
        //SendMessage(Handle, WM_SYSCOMMAND, SC_MAXIMIZE, 0);

        // 取得桌面大小
        SystemParametersInfo(SPI_GETWORKAREA, 0, @ScreenRect, 0);
        Rect := ScreenRect;
        ScreenWidth := ScreenRect.Right - ScreenRect.Left;
        ScreenHeight := ScreenRect.Bottom - ScreenRect.Top;

        m_Param3 := LowerCase(m_Param3);

        If m_Param3 = 'left' Then
        Begin
          Rect.Right := ScreenWidth Div 2;
        End
        Else If m_Param3 = 'right' Then
        Begin
          Rect.Left := ScreenWidth Div 2;
        End
        Else If (m_Param3 = 'top') Or (m_Param3 = 'up') Then
        Begin
          Rect.Bottom := ScreenHeight Div 2;
        End
        Else If (m_Param3 = 'bottom') Or (m_Param3 = 'down') Then
        Begin
          Rect.Top := ScreenHeight Div 2;
        End
        Else If (m_Param3 = 'left+top') Or ((m_Param3 = 'top+left')) Then
        Begin
          Rect.Right  := ScreenWidth Div 2;
          Rect.Bottom := ScreenHeight Div 2;
        End
        Else If (m_Param3 = 'right+top') Or (m_Param3 = 'top+right') Then
        Begin
          Rect.Left := ScreenWidth Div 2;
          Rect.Bottom := ScreenHeight Div 2;
        End
        Else If (m_Param3 = 'left+bottom') Or ((m_Param3 = 'bottom+left')) Then
        Begin
          Rect.Right := ScreenWidth Div 2;
          Rect.Top := ScreenHeight Div 2;
        End
        Else If (m_Param3 = 'right+bottom') Or (m_Param3 = 'bottom+right') Then
        Begin
          Rect.Left := ScreenWidth Div 2;
          Rect.Top  := ScreenHeight Div 2;
        End
        Else
          Try
            List := TStringList.Create;
            SplitString(m_Param3, ',', List);

            // 格式为 “a,b,c,d”，允许丢三落四
            // Left,Top,Width,Height
            For i := 1 To 4 Do
            Begin
              d[i] := -1;

              If List.Count >= i Then
              Begin
                t := StrToIntDef(List.Strings[i - 1], -1);

                // 如果t = -1，说明这是小数，取比例，默认0.2
                If t = -1 Then
                Begin
                  If Odd(i) Then
                    d[i] := Round(ScreenWidth * StrToFloatDef(List.Strings[i - 1], 0.2))
                  Else
                    d[i] := Round(ScreenHeight * StrToFloatDef(List.Strings[i - 1], 0.2));
                End
                Else
                  d[i] := t;
              End;
            End;

            If d[1] >= 0 Then
              Rect.Left := d[1];
            If d[2] >= 0 Then
              Rect.Top := d[2];
            If d[3] >= 0 Then
              Rect.Right := Rect.Left + d[3];
            If d[4] >= 0 Then
              Rect.Bottom := Rect.Top + d[4];

          Finally
            List.Free;
          End;

        MoveWindow(Handle, Rect.Left, Rect.Top, Rect.Right - Rect.Left, Rect.Bottom - Rect.Top, True);
      End;
    End;

    wcTop:
    Begin
      SetWindowPos(Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE);
      SetForegroundWindow(Handle);
    End;

    wcUnTop:
    Begin
      SetWindowPos(Handle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE);
      SetForegroundWindow(Handle);
    End;

    wcHide:
      ShowWindow(Handle, SW_HIDE);

    wcUnHide:
    Begin
      ShowWindow(Handle, SW_SHOW);
      ShowWindow(Handle, SW_SHOW);
      SetForegroundWindow(Handle);
    End;

    wcAlpha:
    Begin
      // 半透明效果
      SetWindowLong(Handle, GWL_EXSTYLE, GetWindowLong(Handle, GWL_EXSTYLE) Or WS_EX_LAYERED);
      SetLayeredWindowAttributes(Handle, 0, StrToIntDef(Param3, 255), LWA_ALPHA);
      SetForegroundWindow(Handle);
    End;

    wcUnAlpha:
    Begin
      // 取消半透明效果
      SetWindowLong(Handle, GWL_EXSTYLE, GetWindowLong(Handle, GWL_EXSTYLE) Or WS_EX_LAYERED);
      SetLayeredWindowAttributes(Handle, 0, 255, LWA_ALPHA);
      SetForegroundWindow(Handle);
    End;

    wcHideTitle:
    Begin
      // 隐藏标题栏
      SetWindowLong(Handle, GWL_STYLE, GetWindowLong(Handle, GWL_STYLE) And (Not WS_CAPTION) And (Not WS_BORDER));
      //SetForegroundWindow(Handle);
      SendMessage(Handle, WM_SYSCOMMAND, SC_MINIMIZE, 0);
      PostMessage(Handle, WM_SYSCOMMAND, SC_RESTORE, 0);
    End;

    wcUnHideTitle:
    Begin
      // 恢复标题栏
      SetWindowLong(Handle, GWL_STYLE, GetWindowLong(Handle, GWL_STYLE) Or WS_CAPTION Or WS_BORDER);
      GetWindowRect(Handle, Rect);
      //SetForegroundWindow(Handle);
      SendMessage(Handle, WM_SYSCOMMAND, SC_MINIMIZE, 0);
      PostMessage(Handle, WM_SYSCOMMAND, SC_RESTORE, 0);
    End;

    wcRound:
    Begin
      // 这个没意义，因为每次窗口刷新，就恢复原型了

      // 圆角矩形窗体
      GetWindowRect(Handle, Rect);
      w := Rect.Right - Rect.Left;
      h := Rect.Bottom - Rect.Top;
      SetWindowRgn(Handle, CreateRoundRectRgn(0, 0, w, h, StrToIntDef(Param3, 0), StrToIntDef(Param3, 0)), True);
    End;

    wcUnRound:
    Begin
      // XP下窗口默认就是上面圆角，下面方形的

      // 取消圆角矩形窗体
      //        GetWindowRect(Handle, Rect);
      //        w := Rect.Right - Rect.Left;
      //        h := Rect.Bottom - Rect.Top;
      //        SetWindowRgn(Handle, CreateRoundRectRgn(0, 0, w, h, 0, 0), True);

      // 用恢复标题栏的方法来恢复窗口
      //SetWindowLong(Handle, GWL_STYLE, GetWindowLong(Handle, GWL_STYLE) or WS_CAPTION);


      SendMessage(Handle, WM_SYSCOMMAND, SC_MINIMIZE, 0);
      PostMessage(Handle, WM_SYSCOMMAND, SC_RESTORE, 0);

    End;

    wcShowOnly:
    Begin
      ShellFunction(sf_MinimizeAll);
      Sleep(300);
      PostMessage(Handle, WM_SYSCOMMAND, SC_RESTORE, 0);
    End;

    wcTest:
    Begin
      //SetWindowLong(Handle, GWL_STYLE, GetWindowLong(Handle, GWL_STYLE) and (not WS_CAPTION));
      //SetWindowLong(Handle, GWL_STYLE, GetWindowLong(Handle, GWL_STYLE) or (WS_CAPTION));
      //GetWindowRect(Handle, Rect);
      //SetWindowRgn(Handle, CreateRoundRectRgn(0, 0, Rect.Right - Rect.Left, Rect.Bottom - Rect.Top, 20, 20), True);
      //GetLastActivePopup()
      //Index := m_WinList.Count mod StrToIntDef(Param3, 0);

      // 取得桌面大小
      SystemParametersInfo(SPI_GETWORKAREA, 0, @Rect, 0);
      //TraceMsg('%d,%d,%d,%d', [Rect.Left, Rect.Right, Rect.Top, Rect.Bottom]);

      MoveWindow(Handle, Rect.Left, Rect.Top, Rect.Right - Rect.Left, Rect.Bottom - Rect.Top, True);
    End;
  End;

  Sleep(50);

  Result := True;
End;

Constructor TWinController.Create;
Begin
  m_WinList := TStringList.Create;
  m_Regex := TRegExpr.Create;
  m_HandleList := TStringList.Create;

  m_WinHandle := 0;

  m_HandleFileName := GetTempDirectory + HANDLE_FILE;
  If FileExists(HANDLE_FILE) Then
    MoveFile(PChar(HANDLE_FILE), PChar(m_HandleFileName));
End;

Destructor TWinController.Destroy;
Begin
  m_WinList.Free;
  m_Regex.Free;
  m_HandleList.Free;
End;

Procedure TWinController.GetForegroundWindowHandle;
Begin
  m_ForegroundWindowHandle := GetForegroundWindow;
End;

Function TWinController.GetParamValue(str, Name: String): String;
Var
  iPos: Integer;
Begin
  Result := '';

  // 如果 str 不是以 Name 开头，则退出
  If Pos(LowerCase(Name), LowerCase(str)) <> 1 Then
    Exit;

  iPos := Pos('=', str);
  If iPos <= 0 Then
    Exit;

  Result := Trim(Copy(str, iPos + 1, Length(str) - iPos));
End;

Function TWinController.GetWinList(str: String; MatchType: TMatchType; MatchRule: TMatchRule): Boolean;
Var
  i: Cardinal;
  Index: Integer;
Begin
  Result := False;

  If MatchType = mtEmpty Then
    Exit;

  m_WinList.Clear;

  If Length(m_WinPropertyList) > 0 Then
  Begin
    If MatchType = mtNext Then
    Begin
      // 先找当前ForegroundWindow的下标
      Index := 0;
      For i := 0 To Length(m_WinPropertyList) - 1 Do
      Begin
        If m_WinPropertyList[i].Handle = m_ForegroundWindowHandle Then
          Break;
      End;

      //TraceMsg('Current = %d, Offset = %s', [i, str]);

      If i = Length(m_WinPropertyList) Then
        i := 0
      Else
        TraceMsg('Current Caption = %s', [m_WinPropertyList[i].Caption]);

      Index := (i + StrToIntDef(str, 1)) Mod Length(m_WinPropertyList);

      // 支持Next后跟负数做参数
      If Index < 0 Then
        Index := Length(m_WinPropertyList) + Index;

      TraceMsg('Next = %d, Caption = %s', [Index, m_WinPropertyList[Index].Caption]);

      Case MatchRule Of
        mrEqual:
          m_WinList.Add(IntToStr(m_WinPropertyList[Index].Handle));

        mrNotEqual:
        Begin
          For i := 0 To Length(m_WinPropertyList) - 1 Do
          Begin
            If i <> Index Then
              m_WinList.Add(IntToStr(m_WinPropertyList[i].Handle));
          End;
        End;
      End;
    End
    Else
      For i := 0 To High(m_WinPropertyList) Do
      Begin
        Case MatchType Of
          mtHandle:
            If IsMatch(IntToStr(m_WinPropertyList[i].Handle), str, MatchRule) Then
              m_WinList.Add(IntToStr(m_WinPropertyList[i].Handle));

          mtCaption:
            If IsMatch(m_WinPropertyList[i].Caption, str, MatchRule) Then
              m_WinList.Add(IntToStr(m_WinPropertyList[i].Handle));

          mtClass:
            If IsMatch(m_WinPropertyList[i].ClassName, str, MatchRule) Then
              m_WinList.Add(IntToStr(m_WinPropertyList[i].Handle));
        End;
      End;
  End;

  //Result := m_WinList.Count > 0;
  Result := True;
End;

Procedure TWinController.RefreshWinPropertyList;
Var
  i: Cardinal;
  str: String;
Begin
  TraceMsg('RefreshWinPropertyList');

  SetLength(m_WinPropertyList, 0);
  EnumWindows(@EnumWindowsInTaskBarFunc, LParam(@m_WinPropertyList));

  str := '';
  If Length(m_WinPropertyList) = 0 Then
    Exit;

  For i := 0 To Length(m_WinPropertyList) - 1 Do
  Begin
    str := Format('[%d] %s%s', [i, m_WinPropertyList[i].Caption, #13#10]);
    TraceMsg(str);
  End;
End;

Function TWinController.HandleCmd(WinCmd: TWinCmdType): Boolean;
Var
  i: Cardinal;
  WinHandle: Integer;
  str: String;
  strFlag: String;
Begin
  Result := False;

  //TraceMsg('HandleCmd = %d', [Ord(m_WinCmd)]);
  Case WinCmd Of
    wcUnkown:
      Exit;

    wcMax,
    wcMin,
    wcRestore,
    wcClose,
    wcMove,
    wcTest:
    Begin
      If m_WinList.Count = 0 Then
        Exit;

      For i := 0 To m_WinList.Count - 1 Do
      Begin
        ControlWindow(StrToInt(m_WinList.Strings[i]), WinCmd);
      End;
    End;

    wcTop,
    wcHide,
    wcAlpha,
    wcHideTitle,
    wcRound:
    Begin
      If m_WinList.Count = 0 Then
        Exit;

      If FileExists(m_HandleFileName) Then
        m_HandleList.LoadFromFile(m_HandleFileName);

      // 兼容老版本
      CompitableWithOldVersion(m_HandleList);

      Case WinCmd Of
        wcTop:
          strFlag := TOP_FLAG;

        wcHide:
          strFlag := HIDE_FLAG;

        wcAlpha:
          strFlag := ALPHA_FLAG;

        wcHideTitle:
          strFlag := TITLE_FLAG;

        wcRound:
          strFlag := ROUND_FLAG;
      End;

      For i := 0 To m_WinList.Count - 1 Do
      Begin
        If ControlWindow(StrToInt(m_WinList.Strings[i]), WinCmd, m_Param3) Then
        Begin
          m_HandleList.Add(strFlag + m_WinList.Strings[i]);
          Sleep(100);
        End;
      End;

      m_HandleList.SaveToFile(m_HandleFileName);
    End;

    wcUnTop,
    wcUnHide,
    wcUnAlpha,
    wcUnHideTitle,
    wcUnRound:
    Begin
      If FileExists(m_HandleFileName) Then
        m_HandleList.LoadFromFile(m_HandleFileName);

      If m_HandleList.Count = 0 Then
        Exit;

      // 兼容老版本
      CompitableWithOldVersion(m_HandleList);

      Case WinCmd Of
        wcUnTop:
          strFlag := TOP_FLAG;

        wcUnHide:
          strFlag := HIDE_FLAG;

        wcUnAlpha:
          strFlag := ALPHA_FLAG;

        wcUnHideTitle:
          strFlag := TITLE_FLAG;

        wcUnRound:
          strFlag := ROUND_FLAG;
      End;

      // 从后向前找
      For i := m_HandleList.Count - 1 Downto 0 Do
      Begin
        str := Trim(m_HandleList.Strings[i]);

        // 不符合的项，直接跳过
        If Pos(strFlag, str) = 0 Then
          Continue;

        // 去除前导项
        str := CutLeftString(str, strFlag);

        // 取得最近的一个
        WinHandle := StrToIntDef(str, 0);

        // 如果Handle错误，则删除此项
        If Not IsWindow(WinHandle) Then
        Begin
          m_HandleList.Delete(i);
          Continue;
        End;

        // 检查是否符合过滤规则，
        Case m_WinParam.MatchType Of
          mtEmpty:
          Begin
            m_HandleList.Delete(i);
            ControlWindow(WinHandle, WinCmd);

            // 找到一个，就到此为止
            Break;
          End;

          mtSingle:
            If WinHandle = m_WinHandle Then
            Begin
              m_HandleList.Delete(i);
              ControlWindow(WinHandle, WinCmd);
            End;

          mtAll:
          Begin
            m_HandleList.Delete(i);
            ControlWindow(WinHandle, WinCmd);
          End;

          mtHandle:
            If IsMatch(IntToStr(WinHandle), m_WinParam.ParamValue, m_WinParam.MatchRule) Then
            Begin
              m_HandleList.Delete(i);
              ControlWindow(WinHandle, WinCmd);
            End;

          mtCaption:
            If IsMatch(GetWinCaption(WinHandle), m_WinParam.ParamValue, m_WinParam.MatchRule) Then
            Begin
              m_HandleList.Delete(i);
              ControlWindow(WinHandle, WinCmd);
            End;

          mtClass:
            If IsMatch(GetWinClassName(WinHandle), m_WinParam.ParamValue, m_WinParam.MatchRule) Then
            Begin
              m_HandleList.Delete(i);
              ControlWindow(WinHandle, WinCmd);
            End;

          mtNext:
          Begin
            m_HandleList.Delete(i);
            ControlWindow(WinHandle, WinCmd);
          End;
        End;
      End;

      m_HandleList.SaveToFile(m_HandleFileName);
    End;

    wcCascade:
      ShellFunction(sf_CascadeWindows);

    wcTileHorizontally:
      ShellFunction(sf_TileHorizontally);

    wcTileVertically:
      ShellFunction(sf_TileVertically);

    wcMinAll:
      ShellFunction(sf_MinimizeAll);

    wcUnMinAll:
      ShellFunction(sf_UndoMinimizeALL);

    wcShowOnly:
    Begin
      If m_WinList.Count = 0 Then
        Exit;

      For i := 0 To m_WinList.Count - 1 Do
      Begin
        ControlWindow(StrToInt(m_WinList.Strings[i]), WinCmd);
      End;
    End;

    wcGetHandle:
    Begin
      // 只接受一个句柄
      If m_WinList.Count <> 1 Then
        Exit;

      Clipboard.AsUnicodeText := m_WinList.Strings[0];
    End;

    wcGetCaption:
    Begin
      // 只接受一个句柄
      If m_WinList.Count <> 1 Then
        Exit;

      Clipboard.AsUnicodeText := GetWinCaption(StrToInt(m_WinList.Strings[0]));
    End;

    wcGetClass:
    Begin
      // 只接受一个句柄
      If m_WinList.Count <> 1 Then
        Exit;

      Clipboard.AsUnicodeText := GetWinClassName(StrToInt(m_WinList.Strings[0]));
    End;
  End;
End;

Function TWinController.HandleCmdList: Boolean;
Var
  i: Cardinal;
Begin
  If Length(m_WinCmdList) > 0 Then
    For i := 0 To Length(m_WinCmdList) - 1 Do
    Begin
      HandleCmd(m_WinCmdList[i]);
    End;
End;

Function TWinController.IsMatch(strText, strPattern: String; MatchRule: TMatchRule): Boolean;
Begin
  Result := False;

  Case MatchRule Of
    mrEqual:
      Result := (strText = strPattern);

    mrNotEqual:
      Result := (strText <> strPattern);

    mrRegexEqual:
    Begin
      m_Regex.Expression := strPattern;
      Try
        Result := m_Regex.Exec(strText);
      Except
        on E: Exception Do
          Result := False;
      End;
    End;

    mrNotRegexEqual:
    Begin
      m_Regex.Expression := strPattern;
      Try
        Result := Not m_Regex.Exec(strText);
      Except
        on E: Exception Do
          Result := False;
      End;
    End;
  End;
End;

Function TWinController.ParseWinParam(str: String; Var WinParam: TWinParam): Boolean;

  Function SplitWinParam(Delimiter: String): Boolean;
  Var
    FlagPos: Integer;
    Len: Integer;
  Begin
    Result := False;

    Len := Length(Delimiter);
    FlagPos := Pos(Delimiter, str);
    If FlagPos <= 0 Then
      Exit;

    WinParam.ParamName := Copy(str, 1, FlagPos - 1);
    WinParam.ParamValue := RemoveQuotation(Copy(str, FlagPos + Len, Length(str) - FlagPos - Len + 1));
    Result := (WinParam.ParamValue <> '');
  End;

Begin
  Result := False;

  If SplitWinParam('!=') Then     // 不等于
  Begin
    WinParam.MatchRule := mrNotEqual;
  End
  Else If SplitWinParam('=') Then // 等于
  Begin
    WinParam.MatchRule := mrEqual;
  End
  Else If SplitWinParam('!~') Then // 不约等于
  Begin
    WinParam.MatchRule := mrNotRegexEqual;
  End
  Else If SplitWinParam('~') Then // 约等于
  Begin
    WinParam.MatchRule := mrRegexEqual;
  End
  Else
    Exit;

  If LowerCase(WinParam.ParamName) = 'handle' Then
    WinParam.MatchType := mtHandle
  Else If LowerCase(WinParam.ParamName) = 'caption' Then
    WinParam.MatchType := mtCaption
  Else If LowerCase(WinParam.ParamName) = 'class' Then
    WinParam.MatchType := mtClass
  Else If LowerCase(WinParam.ParamName) = 'next' Then
    WinParam.MatchType := mtNext;

  Result := True;
End;

Function TWinController.ReadWinCmdList(str: String): Boolean;
Var
  List: TStringList;
  i: Cardinal;
Begin
  Result := True;

  str := LowerCase(str);

  Try
    List := TStringList.Create;
    SplitString(str, '+', List);

    SetLength(m_WinCmdList, 0);

    If List.Count > 0 Then
      For i := 0 To List.Count - 1 Do
      Begin
        // 扩充命令列表
        SetLength(m_WinCmdList, Length(m_WinCmdList) + 1);

        If List.Strings[i] = 'max' Then
          m_WinCmdList[High(m_WinCmdList)] := wcMax
        Else If List.Strings[i] = 'min' Then
          m_WinCmdList[High(m_WinCmdList)] := wcMin
        Else If List.Strings[i] = 'restore' Then
          m_WinCmdList[High(m_WinCmdList)] := wcRestore
        Else If List.Strings[i] = 'close' Then
          m_WinCmdList[High(m_WinCmdList)] := wcClose
        Else If List.Strings[i] = 'move' Then
          m_WinCmdList[High(m_WinCmdList)] := wcMove
        Else If List.Strings[i] = 'top' Then
          m_WinCmdList[High(m_WinCmdList)] := wcTop
        Else If List.Strings[i] = 'untop' Then
          m_WinCmdList[High(m_WinCmdList)] := wcUnTop
        Else If List.Strings[i] = 'hide' Then
          m_WinCmdList[High(m_WinCmdList)] := wcHide
        Else If List.Strings[i] = 'unhide' Then
          m_WinCmdList[High(m_WinCmdList)] := wcUnHide
        Else If List.Strings[i] = 'alpha' Then
          m_WinCmdList[High(m_WinCmdList)] := wcAlpha
        Else If List.Strings[i] = 'unalpha' Then
          m_WinCmdList[High(m_WinCmdList)] := wcUnAlpha
        Else If List.Strings[i] = 'hidetitle' Then
          m_WinCmdList[High(m_WinCmdList)] := wcHideTitle
        Else If List.Strings[i] = 'unhidetitle' Then
          m_WinCmdList[High(m_WinCmdList)] := wcUnHideTitle
        Else If List.Strings[i] = 'round' Then
          m_WinCmdList[High(m_WinCmdList)] := wcRound
        Else If List.Strings[i] = 'unround' Then
          m_WinCmdList[High(m_WinCmdList)] := wcUnRound
        Else If List.Strings[i] = 'cas' Then
          m_WinCmdList[High(m_WinCmdList)] := wcCascade
        Else If List.Strings[i] = 'th' Then
          m_WinCmdList[High(m_WinCmdList)] := wcTileHorizontally
        Else If List.Strings[i] = 'tv' Then
          m_WinCmdList[High(m_WinCmdList)] := wcTileVertically
        Else If List.Strings[i] = 'minall' Then
          m_WinCmdList[High(m_WinCmdList)] := wcMinAll
        Else If List.Strings[i] = 'unminall' Then
          m_WinCmdList[High(m_WinCmdList)] := wcUnMinAll
        Else If List.Strings[i] = 'showonly' Then
          m_WinCmdList[High(m_WinCmdList)] := wcShowOnly
        Else If List.Strings[i] = 'gethandle' Then
          m_WinCmdList[High(m_WinCmdList)] := wcGetHandle
        Else If List.Strings[i] = 'getcaption' Then
          m_WinCmdList[High(m_WinCmdList)] := wcGetCaption
        Else If List.Strings[i] = 'getclass' Then
          m_WinCmdList[High(m_WinCmdList)] := wcGetClass
        Else If List.Strings[i] = 'test' Then
          m_WinCmdList[High(m_WinCmdList)] := wcTest
        Else
        Begin
          //m_WinCmdList[High(m_WinCmdList)] := wcUnkown;
          //Result := False;
        End;
      End;
  Finally
    List.Free;
  End;

{
  // 下面的代码是仅有一个命令时候的，已作废
  if str = 'max' then
    m_WinCmd := wcMax
  else if str = 'min' then
    m_WinCmd := wcMin
  else if str = 'restore' then
    m_WinCmd := wcRestore
  else if str = 'close' then
    m_WinCmd := wcClose
  else if str = 'top' then
    m_WinCmd := wcTop
  else if str = 'untop' then
    m_WinCmd := wcUnTop
  else if str = 'hide' then
    m_WinCmd := wcHide
  else if str = 'unhide' then
    m_WinCmd := wcUnHide
  else if str = 'alpha' then
    m_WinCmd := wcAlpha
  else if str = 'unalpha' then
    m_WinCmd := wcUnAlpha
  else if str = 'hidetitle' then
    m_WinCmd := wcHideTitle
  else if str = 'unhidetitle' then
    m_WinCmd := wcUnHideTitle
  else if str = 'round' then
    m_WinCmd := wcRound
  else if str = 'unround' then
    m_WinCmd := wcUnRound
  else if str = 'cas' then
    m_WinCmd := wcCascade
  else if str = 'th' then
    m_WinCmd := wcTileHorizontally
  else if str = 'tv' then
    m_WinCmd := wcTileHorizontally
  else if str = 'minall' then
    m_WinCmd := wcMinAll
  else if str = 'unminall' then
    m_WinCmd := wcUnMinAll
  else if str = 'showonly' then
    m_WinCmd := wcShowOnly
  else if str = 'gethandle' then
    m_WinCmd := wcGetHandle
  else if str = 'getcaption' then
    m_WinCmd := wcGetCaption
  else if str = 'getclass' then
    m_WinCmd := wcGetClass
  else if str = 'test' then
    m_WinCmd := wcTest
  else
  begin
    m_WinCmd := wcUnkown;
    Result := False;
  end;
}
End;

Function TWinController.ReadWinParam(WinCmd: TWinCmdType; str: String): Boolean;
Var
  i: Cardinal;
  Index: Integer;
Begin
  Result := False;

  // 替换{%c}为剪贴板文字内容
  str := StringReplace(str, CLIP_FLAG, Clipboard.AsUnicodeText, [rfReplaceAll]);

  // 有些命令不需要其他参数
  Case WinCmd Of
    wcUnkown,
    wcCascade,
    wcTileHorizontally,
    wcTileVertically,
    wcMinAll,
    wcUnMinAll:
    Begin
      Result := True;
      Exit;
    End;
  End;

  // 初始化
  m_WinHandle := 0;
  m_WinList.Clear;

  // 参数有如下类型
  // ddd
  // Handle=ddd | Handle!=ddd
  // Caption="xxx" | Caption!="xxx"
  // Class="xxx" | Class!="xxx"
  // ALL

  // 先看看是不是数字
  m_WinHandle := StrToIntDef(str, 0);

  // 如果是数字，则得到 m_WinHandle
  If m_WinHandle > 0 Then
  Begin
    m_WinParam.MatchType := mtSingle;
    m_WinList.Add(IntToStr(m_WinHandle));
    Result := True;
    Exit;
  End;

  // 刷新窗口列表
  RefreshWinPropertyList;

  // 再看是不是 ALL
  If LowerCase(str) = 'all' Then
  Begin
    m_WinParam.MatchType := mtAll;

    If Length(m_WinPropertyList) > 0 Then
      For i := 0 To High(m_WinPropertyList) Do
        m_WinList.Add(IntToStr(m_WinPropertyList[i].Handle));

    Result := True;
    Exit;
  End;

{  // 再看是不是 NEXT
  if LowerCase(str) = 'next' then
  begin
    TraceMsg('Next');

    m_WinParam.MatchType := mtSingle;

    if Length(m_WinPropertyList) > 0 then
    begin
      // 先找当前ActiveWindow的下标
      i := 0;
      for i := 0 to Length(m_WinPropertyList) - 1 do
      begin
        if m_WinPropertyList[i].Handle = m_ForegroundWindowHandle then Break;
      end;

      TraceMsg('Current = %d, Offset = %s', [i, m_Param3]);

      if i = Length(m_WinPropertyList) then
        i := 0
      else
        TraceMsg('Current Caption = %s', [m_WinPropertyList[i].Caption]);

      i := (i + StrToIntDef(m_Param3, 1)) mod Length(m_WinPropertyList);

      // 支持Next后跟负数做参数
      if i < 0 then i := Length(m_WinPropertyList) + i;

      TraceMsg('Next = %d, Caption = %s', [i, m_WinPropertyList[i].Caption]);

      m_WinList.Add(IntToStr(m_WinPropertyList[i].Handle));
      Result := True;
    end;

    Exit;
  end;
}
  // 检查参数类型
  If Not ParseWinParam(str, m_WinParam) Then
    Exit;

  // 对特殊命令 wcShowOnly 的处理，转化为 wcMin 的反面
  If WinCmd = wcShowOnly Then
  Begin
    WinCmd := wcMin;

    Case m_WinParam.MatchRule Of
      mrEqual:
        m_WinParam.MatchRule := mrNotEqual;

      mrNotEqual:
        m_WinParam.MatchRule := mrEqual;

      mrRegexEqual:
        m_WinParam.MatchRule := mrNotRegexEqual;

      mrNotRegexEqual:
        m_WinParam.MatchRule := mrRegexEqual;
    End;
  End;

  If Not GetWinList(m_WinParam.ParamValue, m_WinParam.MatchType, m_WinParam.MatchRule) Then
    Exit;

  Result := True;
End;

Function TWinController.ShellFunction(ShellFunction: TShellFunction): Boolean;
Var
  OleVar: Variant;//OleVariant;
Begin
  Result := False;

  Try
    // 最小化全部窗口
    CoInitialize(nil);
    OleVar := CreateOleObject('Shell.Application');

    Case ShellFunction Of
      sf_MinimizeAll:
        OleVar.MinimizeAll;

      sf_UndoMinimizeALL:
        OleVar.UndoMinimizeALL;

      sf_CascadeWindows:
        OleVar.CascadeWindows;

      sf_TileHorizontally:
        OleVar.TileHorizontally;

      sf_TileVertically:
        OleVar.TileVertically;
    End;
  Finally
    // 每次CoUninitialize都会出现Access Violation问题
    // CoUninitialize;
    OleVar := varNull;
  End;

  Result := True;
End;

Procedure TWinController.Test;
Var
  i: Cardinal;
Begin
  For i := 1 To ParamCount Do
    ShowMessageFmt('Param %d = %s', [i, ParamStr(i)]);
End;

End.
