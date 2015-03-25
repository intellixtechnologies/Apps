object mainFrm: TmainFrm
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  Caption = 'Convert To PDF'
  ClientHeight = 152
  ClientWidth = 339
  Color = 16776176
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'System'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  Position = poDesktopCenter
  PixelsPerInch = 96
  TextHeight = 16
  object SpeedButton1: TSpeedButton
    Left = 296
    Top = 32
    Width = 23
    Height = 24
    Caption = '...'
    OnClick = SpeedButton1Click
  end
  object Label1: TLabel
    Left = 24
    Top = 16
    Width = 72
    Height = 16
    Caption = 'Select File:'
  end
  object Label2: TLabel
    Left = 24
    Top = 64
    Width = 74
    Height = 16
    Caption = 'Output File:'
  end
  object SpeedButton2: TSpeedButton
    Left = 296
    Top = 80
    Width = 23
    Height = 24
    Caption = '...'
    OnClick = SpeedButton2Click
  end
  object Edit1: TEdit
    Left = 24
    Top = 32
    Width = 273
    Height = 24
    TabOrder = 0
  end
  object Button1: TButton
    Left = 24
    Top = 110
    Width = 75
    Height = 25
    Caption = 'Convert'
    TabOrder = 1
    OnClick = Button1Click
  end
  object Edit2: TEdit
    Left = 24
    Top = 80
    Width = 273
    Height = 24
    TabOrder = 2
  end
  object OpenDialog1: TOpenDialog
    DefaultExt = '*.doc'
    Filter = 
      'All Files|*.*|MS Word (*.doc)|*.doc|MS Word(*.docx)|*.docx|Excel' +
      '(*.xls)|*.xls|Excel(*.xlsx)|*.xlsx|PPT(*.ppt)|*.ppt|PPT(*.pptx)|' +
      '*.pptx'
    InitialDir = 'C:\'
    Left = 112
    Top = 112
  end
  object SaveDialog1: TSaveDialog
    DefaultExt = '*.pdf'
    Filter = 'PDF Doc (*.pdf)|*.pdf'
    InitialDir = 'C:\'
    Left = 144
    Top = 112
  end
end
