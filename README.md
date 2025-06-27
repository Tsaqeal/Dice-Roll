# 🎲 Weekend Dice Game – Excel VBA Fun Project

This project started as a fun weekend activity with my daughter when we lost our physical board game dice. I wanted to create a digital replacement, and it turned into a playful little Excel VBA project!

## 📋 How It Started

We began by using a simple Excel formula on the **Dices** worksheet:

- `=RANDBETWEEN(1,6)` in hidden cells to simulate dice rolls
- `XLOOKUP` to match those numbers with Unicode dice icons:
  
  ```excel
  =XLOOKUP(E6, A4:A9, B4:B9)
The dice icons were formatted to be visible, while everything else was in white font to "hide" it.

🎮 Making It Interactive

To make it more fun, I added two buttons:

1. 🎲 Roll Dices – Runs a macro to generate new dice rolls and logs the result

2. ⛔ End Game – Marks the end of a game in the history table

I also added a background image of a board game to make it more thematic.

🧠 How It Works – VBA Macros
🎲 RollDices Macro


```vb
Sub RollDices()
    Dim ws As Worksheet
    Dim historyWs As Worksheet
    Dim historyTbl As ListObject

    Set ws = ThisWorkbook.Sheets("Roll Dice")
    Set historyWs = ThisWorkbook.Sheets("HistoryTable")
    Set historyTbl = historyWs.ListObjects("HistoryTable")

    Dim dice1 As Integer
    Dim dice2 As Integer
    dice1 = WorksheetFunction.RandBetween(1, 6)
    dice2 = WorksheetFunction.RandBetween(1, 6)

    ' Store dice values in hidden cells
    ws.Range("E6").Value = dice1
    ws.Range("F6").Value = dice2

    ' Add to history table
    With historyTbl.ListRows.Add
        .Range(1, 1).Value = Now
        .Range(1, 2).Value = dice1
        .Range(1, 3).Value = dice2

        ' Format timestamp in gray
        .Range(1, 1).Font.Color = RGB(128, 128, 128)

        ' Highlight doubles in green
        If dice1 = dice2 Then
            .Range(1, 2).Font.Color = RGB(0, 128, 0)
            .Range(1, 3).Font.Color = RGB(0, 128, 0)
        End If
    End With
End Sub
```



⛔ EndGame Macro



```vb

Sub EndGame()
    Dim historyWs As Worksheet
    Dim historyTbl As ListObject

    Set historyWs = ThisWorkbook.Sheets("HistoryTable")
    Set historyTbl = historyWs.ListObjects("HistoryTable")

    With historyTbl.ListRows.Add
        .Range(1, 1).Value = "GAME END"
        .Range(1, 1).Font.Color = RGB(255, 0, 0)
        .Range(1, 1).Font.Bold = True

        .Range(1, 2).ClearContents
        .Range(1, 3).ClearContents
    End With
End Sub
```

🧒 Made for Family Fun
This project was originally made just for fun with my daughter, and it turned into a cool way to learn more about Excel VBA. Feel free to fork it and build your own game features!


📦 File Overview
Roll Dice – Main worksheet with buttons and visible dice

Dices – Helper sheet with logic for dice faces

HistoryTable – Table that stores each roll and game end


✅ To Do
- `= Add sound effects for rolling

- `= Add win conditions or scoring


💡 License
MIT License – free to use and adapt.




