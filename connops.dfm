object Form2: TForm2
  Left = 501
  Top = 308
  BorderIcons = []
  BorderStyle = bsSingle
  Caption = #1055#1088#1086#1074#1077#1088#1082#1072' '#1087#1086#1076#1082#1083#1102#1095#1077#1085#1080#1081' '#1054#1055#1057' ...'
  ClientHeight = 298
  ClientWidth = 385
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object ListBox1: TListBox
    Left = 0
    Top = 0
    Width = 385
    Height = 279
    Align = alClient
    ImeName = 'Russian'
    ItemHeight = 13
    TabOrder = 0
  end
  object stat1: TStatusBar
    Left = 0
    Top = 279
    Width = 385
    Height = 19
    Panels = <
      item
        Width = 50
      end>
  end
  object IdIcmpClient1: TIdIcmpClient
    OnReply = IdIcmpClient1Reply
    Left = 24
    Top = 8
  end
  object tmr1: TTimer
    Enabled = False
    OnTimer = tmr1Timer
    Left = 24
    Top = 56
  end
end
