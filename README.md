Basic bypass with encryption,sleep 

added
sleep functions

```
Private Declare PtrSafe Function Sleep Lib "KERNEL32" (ByVal mili As Long) As Long

    Dim res As Long
    Dim t1 As Date
    Dim t2 As Date
    Dim time As Long
    
    t1 = Now()
    Sleep (2000)
    t2 = Now()
    time = DateDiff("s", t1, t2)
    If time < 2 Then
        Exit Function
    End If
```

Encrypter from program.cs 
![image](https://github.com/user-attachments/assets/4fe84bce-e469-4c81-9f70-348b58db7ba7)


Not so basic is vba stomping which uses flexhex
replacing with null bytes
what the GUI editor uses to link the displayed macros
remove this link in the editor could hide our macro
from within the graphical Office VBA editor
![image](https://github.com/user-attachments/assets/82a324f9-ea3c-4feb-ac9b-f8119802ab65)

Do this by inserting zero block
![image](https://github.com/user-attachments/assets/028330e6-1303-4001-9054-06bea26c51f9)

![image](https://github.com/user-attachments/assets/5a08bbc1-93f1-45db-ae30-2f948af47a24)

No longer show macro
![image](https://github.com/user-attachments/assets/6ab1c362-7729-448a-9435-0a6ea27ef6bb)

editing pcode
partially-compressed version of the VBA source code 
![image](https://github.com/user-attachments/assets/cb5a3f61-d377-48e5-9334-7a5d3b5df273)

We need to select all attribute vb name to the remaining byte
![image](https://github.com/user-attachments/assets/fbf3b889-afb1-4aaa-b6ba-a39d5f79d9a0)

![image](https://github.com/user-attachments/assets/cc2786f2-fe61-47f8-b3d2-a8d8790e3fe6)

edit these to zero blocks
![image](https://github.com/user-attachments/assets/f38cdcd0-36cf-4978-be5b-ce4224d3e8d6)

alot better now
![image](https://github.com/user-attachments/assets/7eed2a50-c4e2-4876-85dc-32219f80e61a)

original 
![image](https://github.com/user-attachments/assets/fdbf6641-4ef6-41e6-9ea0-96c57bc12c0f)


