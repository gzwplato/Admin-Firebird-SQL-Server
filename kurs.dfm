object KS1: TKS1
  Left = 546
  Top = 328
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1050#1091#1088#1089' '#1074#1072#1083#1102#1090
  ClientHeight = 295
  ClientWidth = 364
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
  PixelsPerInch = 96
  TextHeight = 13
  object grp2: TGroupBox
    Left = 0
    Top = 0
    Width = 364
    Height = 295
    Align = alClient
    Caption = #1055#1088#1086#1090#1086#1082#1086#1083':'
    TabOrder = 0
    object mmo1: TMemo
      Left = 2
      Top = 15
      Width = 360
      Height = 278
      Align = alClient
      ImeName = 'Russian'
      ScrollBars = ssBoth
      TabOrder = 0
    end
  end
  object IdIcmpClient1: TIdIcmpClient
    Left = 128
    Top = 160
  end
  object IdHTTP1: TIdHTTP
    MaxLineAction = maException
    ReadTimeout = 0
    AllowCookies = True
    ProxyParams.BasicAuthentication = False
    ProxyParams.ProxyPort = 0
    Request.ContentLength = -1
    Request.ContentRangeEnd = 0
    Request.ContentRangeStart = 0
    Request.ContentType = 'text/html'
    Request.Accept = 'text/html, */*'
    Request.BasicAuthentication = False
    Request.UserAgent = 'Mozilla/3.0 (compatible; Indy Library)'
    HTTPOptions = [hoForceEncodeParams]
    Left = 128
    Top = 64
  end
end
