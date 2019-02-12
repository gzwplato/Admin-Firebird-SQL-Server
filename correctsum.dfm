object CorSUM1: TCorSUM1
  Left = 566
  Top = 449
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1042#1074#1077#1076#1080#1090#1077' '#1089#1091#1084#1084#1091':'
  ClientHeight = 99
  ClientWidth = 249
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  Position = poMainFormCenter
  OnActivate = FormActivate
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object stat1: TStatusBar
    Left = 0
    Top = 80
    Width = 249
    Height = 19
    Panels = <
      item
        Text = #1042#1074#1086#1076' '#1082#1086#1088#1088#1077#1082#1090#1080#1088#1091#1102#1097#1077#1081' '#1089#1091#1084#1084#1099' ...'
        Width = 50
      end>
  end
  object grp1: TGroupBox
    Left = 0
    Top = 0
    Width = 249
    Height = 80
    Align = alClient
    TabOrder = 1
    object corsum: TEdit
      Left = 8
      Top = 17
      Width = 233
      Height = 37
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGreen
      Font.Height = -24
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ImeName = 'Russian'
      ParentFont = False
      TabOrder = 0
      OnChange = corsumChange
    end
    object minus: TCheckBox
      Left = 8
      Top = 57
      Width = 53
      Height = 17
      Caption = #1052#1080#1085#1091#1089
      TabOrder = 1
      OnClick = minusClick
    end
    object plus: TCheckBox
      Left = 183
      Top = 57
      Width = 51
      Height = 17
      Caption = #1055#1083#1102#1089
      TabOrder = 2
      OnClick = plusClick
    end
    object ext: TCheckBox
      Left = 92
      Top = 57
      Width = 53
      Height = 17
      Caption = #1042#1099#1093#1086#1076
      TabOrder = 3
      OnClick = extClick
    end
  end
end
