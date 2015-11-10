object DataModule2: TDataModule2
  OldCreateOrder = False
  Left = 447
  Top = 201
  Height = 212
  Width = 876
  object DataSource1: TDataSource
    DataSet = ADOTable1
    Left = 8
    Top = 8
  end
  object ADOTable1: TADOTable
    Connection = ADOConnection1
    CursorType = ctStatic
    TableName = 'produit'
    Left = 160
    Top = 8
  end
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=BDD_moh.accdb;Pers' +
      'ist Security Info=False;'
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.ACE.OLEDB.12.0'
    Left = 88
    Top = 8
  end
  object DataSource2: TDataSource
    DataSet = ADOTable2
    Left = 16
    Top = 64
  end
  object ADOConnection2: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=BDD_moh.accdb;Pers' +
      'ist Security Info=False;'
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.ACE.OLEDB.12.0'
    Left = 88
    Top = 64
  end
  object ADOTable2: TADOTable
    Connection = ADOConnection2
    CursorType = ctStatic
    TableName = 'facture'
    Left = 160
    Top = 64
  end
  object ADOQuery_rech_produit_date: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    DataSource = DataSource1
    Parameters = <
      item
        Name = 'date'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      'SELECT * '
      'FROM produit'
      'WHERE date_ajout =:date ;')
    Left = 264
    Top = 8
  end
  object ADOQuery_rech_modif_date: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    DataSource = DataSource1
    Parameters = <
      item
        Name = 'date'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      'SELECT * '
      'FROM produit '
      'WHERE date_modif =:date ;')
    Left = 408
    Top = 8
  end
  object ADOQuery_date_semaine: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    DataSource = DataSource1
    Parameters = <
      item
        Name = 'date1'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end
      item
        Name = 'date2'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      'SELECT * '
      'FROM produit '
      'WHERE date_ajout BETWEEN :date1 AND :date2 ; ')
    Left = 544
    Top = 8
  end
  object ADOQuery_qte_critique: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    DataSource = DataSource1
    Parameters = <>
    SQL.Strings = (
      'SELECT * '
      'FROM produit '
      'WHERE quantite_produit<20;')
    Left = 672
    Top = 8
  end
  object ADOQuery_rech_vente: TADOQuery
    Connection = ADOConnection2
    CursorType = ctStatic
    DataSource = DataSource2
    Parameters = <
      item
        Name = 'date'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      'SELECT * '
      'FROM facture '
      'WHERE date_facture=:date;')
    Left = 264
    Top = 64
  end
  object DataSource3: TDataSource
    DataSet = ADOTable3
    Left = 24
    Top = 120
  end
  object ADOTable3: TADOTable
    Connection = ADOConnection3
    CursorType = ctStatic
    TableName = 'num_facture'
    Left = 168
    Top = 120
  end
  object ADOConnection3: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=BDD_moh.accdb;Pers' +
      'ist Security Info=False;'
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.ACE.OLEDB.12.0'
    Left = 88
    Top = 120
  end
  object ADOQuery_vente_date: TADOQuery
    Connection = ADOConnection2
    CursorType = ctStatic
    DataSource = DataSource2
    Parameters = <
      item
        Name = 'date1'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end
      item
        Name = 'date2'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      'SELECT * '
      'FROM facture'
      'WHERE date_facture BETWEEN :date1 AND :date2 ; ')
    Left = 384
    Top = 64
  end
end
