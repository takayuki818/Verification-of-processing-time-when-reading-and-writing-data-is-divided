Attribute VB_Name = "Module1"
Option Explicit
Sub �A������()
    Dim ������ As Long
    With Sheets("�z�񕪊��e�X�g")
        For ������ = 1 To 10
            If Int(100000 / ������) = 100000 / ������ Then
                .Range("������") = ������
                Call �z�񕪊��ǂݏ�������
            End If
        Next
    End With
    MsgBox "���؊���"
End Sub
Sub �z�񕪊��ǂݏ�������()
    Dim ������ As Long, �� As Long, �s As Long, �� As Long, �l As Long, �i�� As Long
    Dim �n�� As Date
    With Sheets("�z�񕪊��e�X�g")
        ������ = .Range("������")
        If Int(100000 / ������) <> 100000 / ������ Then
            MsgBox "���Z�̗]�肪�����Ă��܂�" & vbCrLf & "10���������ė]��̏o�Ȃ�����������͂��Ă�������"
            Exit Sub
        End If
        �n�� = Now
        ���s��.Show vbModeless
        Call ���s��.�v���O���X�o�[�X�V(0)
        For �� = 1 To ������
            ReDim �z��(1 To 100000 / ������, 1 To 100)
            For �s = 1 To 100000 / ������
                For �� = 1 To 100
                    �l = �l + 1
                    �z��(�s, ��) = �l
                Next
                If Int((�s + (�� - 1) * 100000 / ������) / 100000 * 100) - �i�� >= 10 Then
                    �i�� = Int((�s + (�� - 1) * 100000 / ������) / 100000 * 100)
                    Call ���s��.�v���O���X�o�[�X�V(�i��)
                End If
            Next
            Range(.Cells(7 + 100000 / ������ * (�� - 1), 4), .Cells(7 + 100000 / ������ * �� - 1, 103)) = �z��
            Erase �z��
        Next
        .Cells(7, 4).Resize(100000, 100).ClearContents
        .Range("��������") = (Now - �n��) * 24 * 60 * 60
        .Cells(Rows.Count, 1).End(xlUp).Offset(1, 0) = ������
        .Cells(Rows.Count, 1).End(xlUp).Offset(0, 1) = .Range("��������")
        Unload ���s��
    End With
End Sub
