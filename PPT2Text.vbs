Option Explicit

' ファイル名を実行時パラメータから取得する
Dim pptFilename
pptFilename = WScript.Arguments.Item(0)

Dim oApp
Set oApp = CreateObject("PowerPoint.Application")
oApp.Visible = True

oApp.Presentations.Open(pptFilename)

WScript.echo "<?xml version='1.0' encoding='Shift_JIS' ?>" & vbCrLf
WScript.echo "<Slides>" & vbCrLf

' 全スライドに対して処理を行う
Dim pSlide
For Each pSlide In oApp.ActiveWindow.Parent.Slides
        WScript.echo "<Slide><SlideNumber value='" & pSlide.SlideNumber & "'/>" & vbCrLf
        WScript.echo "<SlideBody><![CDATA[" & vbCrLf
        
        ' スライドのテキストを全部表示する
        Dim pShape
        For Each pShape In pSlide.Shapes
                If pShape.HasTextFrame Then
                        If pShape.TextFrame.HasText Then
                                With pShape.TextFrame.TextRange
                                        WScript.echo CleanChar(.Text)
                                End With
                        End If
                End If
        Next
        
        ' ノートのテキストを全部表示する
        For Each pShape In pSlide.NotesPage.Shapes
                If pShape.HasTextFrame Then
                        If pShape.TextFrame.HasText Then
                                With pShape.TextFrame.TextRange
                                        WScript.echo CleanChar(.Text)
                                End With
                        End If
                End If
        Next
        
        WScript.echo "]]></SlideBody>" & vbCrLf
        WScript.echo "</Slide>" & vbCrLf
Next

WScript.echo "</Slides>"

oApp.ActivePresentation.Application.Quit ' PowerPointを終了する

Set oApp = Nothing


' http://www.tsware.jp/tips/tips_406.htm
' 上記サイトから引用
Private Function CleanChar(ByVal strData)
'引数の文字列から制御コードを除去した文字列を返す

  Dim strRet
  Dim strCurChar
  Dim iintLoop

  strRet = ""
  For iintLoop = 1 To Len(strData)
    strCurChar = Mid(strData, iintLoop, 1)
    If Asc(strCurChar) < 0 Or Asc(strCurChar) >= 32 Then
      '漢字のAscの返り値はマイナスに留意
      strRet = strRet & strCurChar
    End If
  Next

  CleanChar = strRet

End Function