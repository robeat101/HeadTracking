   M S B A M L   `   `   `  ����'  #Kinecting for Windows - Expressions     K4W.Expressions.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   $	��720��5   	   $
��1400��6     � Icon$  Resources/icon.png��6#    � Title$7 1Kinecting for Windows - First look at Expressions��5   	   .��5      � -   5         �Name$ 
LayoutRoot��5       � RowDefinitions 5   
   �� 5      $��120px �6   �� 5	      $��2 �6   �� 5
      .+�5   
   �� $
��Left=�6   $��10q�60    �� Orientation$ 
Horizontal=�6<   .+�5      �� $��Resources/k4w.png��6   $��Center=�6/   $��Center=�6L   $	��100��6g   �� 5      $:��4Second Generation Kinect - First look at Expressions��6   $��Center=�6T   $��20(�6o   $
��Bold"�6}   $��20,0,0,0q�6�   ���	#FF2A2A2A6�   5      �� 5   
   -   5         ��Name$ 	Seperator��5      
���   �$��1��64   � 5   
   -   5        $ Content��5      $��2��6     � ColumnDefinitions 5      �� 5      $	��960 �6#   �� 5      �� 5       5      �� 5      �� 5      �� 5       �� 5!      �� 5"      �� 5#      �� 5$      �� 5%      .+�5)      �� -   5         ��Name$ CameraImage��5)      $	��960��6)   $	��540��65   $��8��6B    J� RenderTransformOrigin$ 0.5,0.5&�6S   z�5*      V� 5+      .~�5,      �� 	 �� ScaleY$	 1Q�6)   
 �� ScaleX$
 -1Q�64   �� 5-       �� AngleY$ 0Q�6(    �� AngleX$ 0Q�63   �� 5.       �� Angle$ 0Q�6*   T� 5/      50      52      �� 55      $��1��6   $��Happy��6$   $
��Bold"�64   $��20(�6F   $��Center=�6T   $��Center=�6x   �� 56      $��1��6   $��Engaged��6$   $��1��66   $
��Bold"�6C   $��20(�6U   $��Center=�6c   $��Center=�6�   �� 57      $��1��6   $��Wearing Glasses��6$   $��2��6>   $
��Bold"�6K   $��20(�6]   $��Center=�6k   $��Center=�6�   �� 58      $��1��6   $��Left Eye Closed��6$   $��3��6>   $
��Bold"�6K   $��20(�6]   $��Center=�6k   $��Center=�6�   �� 59      $��1��6   $��Right Eye Closed��6$   $��4��6?   $
��Bold"�6L   $��20(�6^   $��Center=�6l   $��Center=�6�   �� 5:      $��1��6   $��
Mouth Open��6$   $��5��69   $
��Bold"�6F   $��20(�6X   $��Center=�6f   $��Center=�6�   �� 5;      $��1��6   $��Mouth Moved��6$   $��6��6:   $
��Bold"�6G   $��20(�6Y   $��Center=�6g   $��Center=�6�   �� 5<      $��1��6   $��Looking Away��6$   $��7��6;   $
��Bold"�6H   $��20(�6Z   $��Center=�6h   $��Center=�6�   �� 5?      -   5         ��Name$ HappyResult��5?      $��2��6-   $��20(�6=   $��Center=�6K   $��Center=�6h   $��No face tracked��6�   �� 5@      -   5        $ EngagedResult��5@      $��2��6/   $��1��6?   $��20(�6L   $��Center=�6Z   $��Center=�6w   $��No face tracked��6�   �� 5A      -   5        $ GlassesResult��5A      $��2��6/   $��2��6?   $��20(�6L   $��Center=�6Z   $��Center=�6w   $��No face tracked��6�   �� 5B      -   5        $ LeftEyeResult��5B      $��2��6/   $��3��6?   $��20(�6L   $��Center=�6Z   $��Center=�6w   $��No face tracked��6�   �� 5C      -	   5        $ RightEyeResult��5C      $��2��60   $��4��6@   $��20(�6M   $��Center=�6[   $��Center=�6x   $��No face tracked��6�   �� 5D      -
   5        $ MouthOpenResult��5D      $��2��61   $��5��6A   $��20(�6N   $��Center=�6\   $��Center=�6y   $��No face tracked��6�   �� 5E      -   5        $ MouthMovedResult��5E      $��2��62   $��6��6B   $��20(�6O   $��Center=�6]   $��Center=�6z   $��No face tracked��6�   �� 5F      -   5        $ LookingAwayResult��5F      $��2��63   $��7��6C   $��20(�6P   $��Center=�6^   $��Center=�6{   $��No face tracked��6�   5G      5H      5I      