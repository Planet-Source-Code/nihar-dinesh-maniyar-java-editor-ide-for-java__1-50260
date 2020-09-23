Attribute VB_Name = "Module2"
'Function to make a window as a parent
Public Declare Function SetParent Lib "user32" (ByVal hwndchild As Long, ByVal hwndnewparent As Long) As Long
