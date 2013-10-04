Option Explicit

' �t�@�C���������s���p�����[�^����擾����
Dim pptFilename
pptFilename = WScript.Arguments.Item(0)

Dim oApp
Set oApp = CreateObject("PowerPoint.Application")
oApp.Visible = True

oApp.Presentations.Open(pptFilename)

WScript.echo "<?xml version='1.0' encoding='Shift_JIS' ?>" & vbCrLf
WScript.echo "<Slides>" & vbCrLf

' �S�X���C�h�ɑ΂��ď������s��
Dim pSlide
For Each pSlide In oApp.ActiveWindow.Parent.Slides
        WScript.echo "<Slide><SlideNumber value='" & pSlide.SlideNumber & "'/>" & vbCrLf
        WScript.echo "<SlideBody><![CDATA[" & vbCrLf
        
        ' �X���C�h�̃e�L�X�g��S���\������
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
        
        ' �m�[�g�̃e�L�X�g��S���\������
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

oApp.ActivePresentation.Application.Quit ' PowerPoint���I������

Set oApp = Nothing


' http://www.tsware.jp/tips/tips_406.htm
' ��L�T�C�g������p
Private Function CleanChar(ByVal strData)
'�����̕����񂩂琧��R�[�h�����������������Ԃ�

  Dim strRet
  Dim strCurChar
  Dim iintLoop

  strRet = ""
  For iintLoop = 1 To Len(strData)
    strCurChar = Mid(strData, iintLoop, 1)
    If Asc(strCurChar) < 0 Or Asc(strCurChar) >= 32 Then
      '������Asc�̕Ԃ�l�̓}�C�i�X�ɗ���
      strRet = strRet & strCurChar
    End If
  Next

  CleanChar = strRet

End Function