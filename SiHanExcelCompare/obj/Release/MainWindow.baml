   M S B A M L   `   `   `  ���� L  HSiHanExcelCompare, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null5clr-namespace:SiHanExcelCompareSiHanExcelCompare  "    SiHanExcelCompare.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     1localclr-namespace:SiHanExcelCompare       � WindowStartupLocation$  CenterScreen=� � 
ResizeMode$ NoResize=� � Title$ 思晗对比工具��$	��850��$
��1400��$	��850��$
��1400��$	��850��$
��1400��.��� -   $��grid1��$��24q� � ColumnDefinitions �� $	��670 ��� $	��670 � � RowDefinitions �� $	��200 ��� $��29 ��� $	��490 ��� $	��53* �.+��� $��0��$��0��$��0,0,4,28q�$��2��.��� -   $��
sourceGrid��$��8q� �� $��49 ��� $	��231 ��� $��84 ��� $	��13* ��� $	��69* � �� $��43 ��� $
��128* �.+��� $��0��$��0��$��Right=�$��Center=�$��0,8,0,9q�$��26��$��46��.��	原表：�� -    ��Name$ sourceFileName_text��$��0��$��1��$��Center=�$��22��$��	0,10,0,11q��� -    ��Name$ sourceBrowser_btn��$��0��$��2��$��Center=�$��22��$��Center=�$��	0,10,0,11q�.��浏览�� -   $��sourceSheetList��$��1��$��3��$��0,0,8,0q� �� View �� P��� $�� �� �� Width$ 20��	 �� DisplayMemberBinding	 �� 
 �� Path$
 Index��� $��表单��$	 150��	 �� $

 Name��� $��表头行号��$ 80�� �� CellTemplate �� .R��� $ sourceSheetLineNum_text��-   $��80��$��80������ $
 LineNum��� $��选中��$ 50��	 �� $
 Selected�����  �� 
TargetType" ���-   �� -   $��sourceHeaderList��$��1��$��3��$��3��$��8,0,0,0q� �� P��� $	��列��$ 50��	 �� $
 Column��� $��	表头名��$	 100��	 �� $

 Name��� $��0��$��1��$��4,0,0,28q�$��2��.��� -	   $��
targetGrid��$��8q� �� $��49 ��� $	��231 ��� $��84 ��� $	��13* ��� $	��79* � �� $��43 ��� $
��128* �.+��� $��0��$��0��$��Right=�$��Center=�$��0,8,0,9q�$��26��$��46��$	��0,5q�.��目标表：�� -
   $ targetFileName_text��$��0��$��1��$��Center=�$��22��$��	0,10,0,11q��� -   $ targetBrowser_btn��$��0��$��2��$��Center=�$��22��$��Center=�$��	0,10,0,11q�.��浏览�� -   $��targetSheetList��$��1��$��3�� �� P��� $�� ��$ 20��	 �� $
 Index��� $	��表��$	 150��	 �� $

 Name��� $��表头行号��$ 80�� �� .R��� $ targetSheetLineNum_text��-   $��80��$��80������ $
 LineNum��� $��选中��$ 50��	 �� $
 Selected����� " ���-   �� -   $��targetHeaderList��$��1��$��3��$��3��$��8,0,0,0q� �� P��� $	��列��$ 50�� �� .R��� $# targetHeaderTargetColumn_text��-   $��50��$��50��$��0������ $
 TargetColumn��� $��	表头名��$	 250��	 �� $

 Name�� $��1��$��2��$��	0,1,0,488q�$��2�� �� $��96 ��� $	��574 ��� $
��467* ��� $
��108* ��� $	��97* �.+��� -   $ 
export_btn��$��0��$��0��.��导出结果�� -   $ export_btn_all��$��0��$��1��$
��Left=�$	��102��$��30,0,0,0q�.��导出原表�� -   $ Compare_btn��$��0��$��4��$
��Left=�$��79��$��2,1,1,1q�$��8,0,0,0q�.��对比�� -    ��Name$ 
showAll_ck��$��0��$��3��$��Center=�$��LeftToRight=� d� 	IsChecked$
 TrueY�.��显示全部�� $��2��$��2��$��0,2,0,3q�$��2��.+�� -   $��
scrollGrid��$��8q� �� $	��660 ��� $	��660 �.+�&   System.Windows.Controls.DataGrid  -   $��sourceTableData��$��0��$��0��$��0,0,0,0q�$	��650��$	��650��$	��470��$	��470��   GridLinesVisibility$ 
Horizontal=�   
IsReadOnly .   -   $��targetTableData��$��0��$��1��$��10,0,0,0q�$	��650��$	��650��$	��470��$	��470��$ 
Horizontal=�   HorizontalGridLinesBrush
 �����   VerticalGridLinesBrush
 ����� . � $��4��$��2�� �� $	��669 ��� $	��189 �.+��� -    ��Name$ sourceResult_label��$��0��$��0��$��Center=�$
��Left=��� -   $ targetResult_label��$��0��$��1��$��Center=�$
��Left=�