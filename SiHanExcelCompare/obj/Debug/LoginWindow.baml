   M S B A M L   `   `   `  ����L  HSiHanExcelCompare, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null5clr-namespace:SiHanExcelCompareSiHanExcelCompare  5   	   #    SiHanExcelCompare.LoginWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   1localclr-namespace:SiHanExcelCompare     5   	     � WindowStartupLocation$  CenterScreen=�5   	    � Title$ 登录��5	   	   $	��248��6   $	��400��6!   .��5
      � -   5        $��grid1��5
      $��24q�6    � ColumnDefinitions 5   
   �� 5      $��50 �6   �� 5      $��* �6    � RowDefinitions 5   
   �� 5      $��50 �6   �� 5      $��40 �6   �� 5      $��40 �6   �� 5      $��* �6   .+�5   
   �� $��2��6   $��0��6$   $��Center=�61   $	��Top=�6N   $��22(�6f   $
��Bold"�6t   $��Italic$�6�   .��6�   思晗对比工具6�   �� 5   
   $��1��6   $��0��6   $��Center=�6-   $��Right=�6H   .��6d   	账号：6i   �� 5   
   -   5         ��Name$ account_text��5      $��1��6(   $��1��65   $��20(�6F   $�� ��6T   $
��Left=�6\   $��Center=�6w   $	��294��6�   �� 5   
   $��2��6   $��0��6   $��Right=�6-   $��Center=�6I   .��6d   	密码：6i   G� 5   
   -   5         G�Name$ password_text��5      $��2��6.   $��1��6;   $��20(�6L   $
��Left=�6Z   $��Center=�6u   $	��294��6�   4  .System.Windows.Automation.AutomationProperties   HelpText$ password��6�   �� 5   
   -   5         ��Name$ 	Login_btn��5      $��3��6$   $��2��61   $��Center=�6E   $��Center=�6b   $	��100��6}   $��30��6�   .��6�   登录6�   5      5      