# การสร้างมาโครใน Microsoft Word / Creating Macros in Microsoft Word

## ภาษาไทย

### บทนำ
มาโคร (Macro) เป็นเครื่องมือที่ช่วยในการทำงานซ้ำ ๆ อย่างอัตโนมัติใน Microsoft Word โดยการบันทึกชุดคำสั่งที่เราทำบ่อย ๆ แล้วนำมาใช้ใหม่เมื่อเราต้องการ ทำให้สามารถประหยัดเวลาและเพิ่มประสิทธิภาพในการทำงานได้

### ขั้นตอนการสร้างมาโคร
1. **เปิด Microsoft Word**
2. **ไปที่แท็บ "View" (มุมมอง)**
   - เลือก "Macros" (มาโคร) > "Record Macro" (บันทึกมาโคร)
3. **ตั้งชื่อมาโคร**
   - กรอกชื่อมาโครในช่อง "Macro name" (ชื่อมาโคร)
   - เลือกว่าจะเก็บมาโครไว้ใน "All Documents (Normal.dotm)" (เอกสารทั้งหมด) หรือเอกสารปัจจุบัน
4. **เลือกว่าจะผูกมาโครกับปุ่มหรือคีย์ลัด**
   - คลิกที่ "Button" (ปุ่ม) หรือ "Keyboard" (คีย์บอร์ด) เพื่อตั้งค่า
5. **ใส่โค้ด VBA ที่ต้องการ** 
   - เปิดหน้าต่าง Visual Basic for Applications (VBA) โดยไปที่ "View" > "Macros" > "View Macros" แล้วคลิก "Edit" 
   - ใส่โค้ด VBA ในหน้าต่างที่เปิดขึ้น
   - ตัวอย่าง:
```vba
Sub arabic2thai()
  For i = 0 To 9
    With Selection.Find
      .Text = Chr(48 + i)
      .Replacement.Text = Chr(240 + i)
      .Wrap = wdFindContinue
    End With
  Selection.Find.Execute Replace:=wdReplaceAll
  Next
End Sub

Sub thai2arabic()
  For i = 0 To 9
    With Selection.Find
      .Text = Chr(240 + i)
      .Replacement.Text = Chr(48 + i)
      .Wrap = wdFindContinue
    End With
  Selection.Find.Execute Replace:=wdReplaceAll
  Next
End Sub
```

### การใช้งานมาโคร
1. **ไปที่แท็บ "View" (มุมมอง)**
   - เลือก "Macros" > "View Macros" (ดูมาโคร)
2. **เลือกมาโครที่ต้องการใช้งาน**
   - เลือกชื่อมาโครที่ต้องการใช้งาน
   - คลิก "Run" (เรียกใช้งาน)

# Introduction
## A macro in Microsoft Word is a tool that allows you to automate repetitive tasks by recording a series of commands that you perform frequently. By using macros, you can save time and enhance your productivity.

### Steps to Create a Macro
1. **Open Microsoft Word**
2. **Go to the "View" tab**
  - Select "Macros" > "Record Macro"
3. **Name your macro**
  - Enter a name for your macro in the "Macro name" field
  - Choose whether to store the macro in "All Documents (Normal.dotm)" or the current document
4. **Choose to assign the macro to a button or keyboard shortcut**
  - Click "Button" or "Keyboard" to configure
5. **Enter the desired VBA code**
  - Open the Visual Basic for Applications (VBA) window by going to "View" > "Macros" > "View Macros" and clicking "Edit"
  - Enter the VBA code in the window that opens
  - Example:
```vba
Sub arabic2thai()
  For i = 0 To 9
    With Selection.Find
      .Text = Chr(48 + i)
      .Replacement.Text = Chr(240 + i)
      .Wrap = wdFindContinue
    End With
  Selection.Find.Execute Replace:=wdReplaceAll
  Next
End Sub

Sub thai2arabic()
  For i = 0 To 9
    With Selection.Find
      .Text = Chr(240 + i)
      .Replacement.Text = Chr(48 + i)
      .Wrap = wdFindContinue
    End With
  Selection.Find.Execute Replace:=wdReplaceAll
  Next
End Sub
```
***Using a Macro***
1. **Go to the "View" tab**
- Select "Macros" > "View Macros"
2. **Choose the macro you want to run**
- Select the macro name you want to run
- Click "Run"
