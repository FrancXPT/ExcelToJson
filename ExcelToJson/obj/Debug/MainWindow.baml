   M S B A M L   `   `   `  ����J  FExcel Converter, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null)clr-namespace:ExcelToJsonExcelToJson  5   	       ExcelToJson.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   +localclr-namespace:ExcelToJson     5   	      MaterialDesignBody#���   5   	     MaterialDesignPaper#���  5   	     � Title$#  Convertion Du Base des Donner��5
   	   $��353.334��6/   $��343.667��6@    � 
ResizeMode$ NoResize=�6P    � WindowStartupLocation$ CenterScreen=�6f    � WindowStyle$
 None=�6�   .��5      � .+�5   
   �� -   5         ��Name$ chossedFiletxt��5      $��Right=�6*   $��27��6F    �� TextWrapping$
 Wrap=�6R   $��Bottom=�6f   $	��262��6�   $��
0,0,49,214q�6�   �� 5   
   -   5         ��Name$ 	ChosseBtn��5      $
��....��6$   $
��Left=�63   $��
284,48,0,0q�6N   $	��Top=�6b   $��44��6z   6�   �� 5   
   $��Choisir Un fichier Excel��6   $
��Left=�63   $	��Top=�6N   $��	29,46,0,0q�6f   $	��158��6y   9� 5   
   -   5         9�Name$ Expande��5      $��Expander��6$   $
��Left=�66   $	��146��6Q   $��	9,150,0,0q�6^   $	��Top=�6q   $	��315��6�    � Cursor$
 Hand��6�   .��5      � $
��Left=�6   $��87��6.   $��2,0,0,0q�6:   $	��311��6K   	 � ColumnDefinitions	 5      �� 5      $
��135* �6'   �� 5      $
��176* �6'   .+�5      �� -   5        $ SaveLocationBtn��5      $
��....��62   $
��Left=�6A   $��
119,46,0,0q�6\   $	��Top=�6p   $��44��6�   $��1��6�   6�   �� 5      -   5        $ SaveLocationtxt��5      $��Right=�63   $��27��6O   $
 Wrap=�6[   $��Bottom=�6o   $	��276��6�   $��	0,0,16,50q�6�   $��2��6�   5      5      �� 5   
   -   5        
 ��Name$
 CloseBtn��5      $
��Left=�6"   $��31��6=   $��	308,4,0,0q�6I   $	��Top=�6\   $��31��6t   $��icons8_Cancel_64.png��6   $
 Hand��6�   6�   �� 5   
   -   5        $ 
ConvertBtn��5      $��	Convertir��6%   $
��Left=�69   $��210,310,0,0q�6T   $	��Top=�6i   $	��104��6�   6�   �� 5   
   -   5         ��Name$ progressPanel��5      $
��Left=�6,   $	��257��6G   $��
12,43,0,53q�6T   $��Center=�6h   $	��320��6�   ���	#FF303030$y�Hidden=�6�   .+�5      � $��63��6   $��Center=�6&   $	��286��6C   $��
18,50,16,0q�6O    � IsIndeterminate .  � LargeChange$ 5Q�6z    � Minimum$ 10Q�6�   5      5      5      