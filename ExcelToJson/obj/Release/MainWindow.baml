   M S B A M L   `   `   `  ���� J  FExcel Converter, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null)clr-namespace:ExcelToJsonExcelToJson      ExcelToJson.MainWindow   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     +localclr-namespace:ExcelToJson        MaterialDesignBody#���     MaterialDesignPaper#���    � Title$#  Convertion Du Base des Donner��$��353.334��$��343.667�� � 
ResizeMode$ NoResize=� � WindowStartupLocation$ CenterScreen=� � WindowStyle$
 None=�.��� .+��� -    ��Name$ chossedFiletxt��$��Right=�$��27�� �� TextWrapping$
 Wrap=�$��Bottom=�$	��262��$��
0,0,49,214q��� -    ��Name$ 	ChosseBtn��$
��....��$
��Left=�$��
284,48,0,0q�$	��Top=�$��44���� $��Choisir Un fichier Excel��$
��Left=�$	��Top=�$��	29,46,0,0q�$	��158��9� -    9�Name$ Expande��$��Expander��$
��Left=�$	��146��$��	9,150,0,0q�$	��Top=�$	��315�� � Cursor$
 Hand��.��� $
��Left=�$��87��$��2,0,0,0q�$	��311��	 � ColumnDefinitions	 �� $
��135* ��� $
��176* �.+��� -   $ SaveLocationBtn��$
��....��$
��Left=�$��
119,46,0,0q�$	��Top=�$��44��$��1���� -   $ SaveLocationtxt��$��Right=�$��27��$
 Wrap=�$��Bottom=�$	��276��$��	0,0,16,50q�$��2���� -   
 ��Name$
 CloseBtn��$
��Left=�$��31��$��	308,4,0,0q�$	��Top=�$��31��$��icons8_Cancel_64.png��$
 Hand���� -   $ 
ConvertBtn��$��	Convertir��$
��Left=�$��210,310,0,0q�$	��Top=�$	��104���� -    ��Name$ progressPanel��$
��Left=�$	��257��$��
12,43,0,53q�$��Center=�$	��320�����	#FF303030$y�Hidden=�.+�� $��63��$��Center=�$	��286��$��
18,50,16,0q� � IsIndeterminate .  � LargeChange$ 5Q� � Minimum$ 10Q�