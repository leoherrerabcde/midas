Attribute VB_Name = "modBackGroundMessageProcess"

Function Parse_Header_Message(lsMsg() As String) As Message_Header_Const

Dim i           As Message_Header_Const

    For i = Message_Header_Const.MSG_RUNXLS To Message_Header_Const.MSG_END_PROJECT
        If lsMsg(0) = GV_Msg_Header(i) Then
            Parse_Header_Message = i
            Exit Function
        End If
    Next
    Parse_Header_Message = MSG_ERROR

End Function

Function Verify_Length_Msg(lsMsg() As String, IndexMsg As Message_Header_Const) As Boolean

    If IndexMsg = MSG_ERROR Then
        Verify_Length_Msg = False
        Exit Function
    End If
    If GV_Msg_Len(IndexMsg) = UBound(lsMsg) Then
        Verify_Length_Msg = True
    Else
        Verify_Length_Msg = False
    End If
    
End Function

Sub DiscardBackGroundProcess()

    If BackGroundProjectList.ListIndex >= 0 Then
        With BackGroundProjectList
            With .ProjectList(.ListIndex)
                .ProjectStarted = False
            End With
            .ListIndex = -1
        End With
    End If
    
End Sub

Sub AddBgProject(TrVw As TreeView, BgPjt As RemoteProject, IndexLastNode As Long)

Dim lvNode          As Node

    If IndexLastNode Then
        Set lvNode = TrVw.Nodes.Add(IndexLastNode, tvwNext, , BgPjt.ProjectName)
    Else
        Set lvNode = TrVw.Nodes.Add(, tvwFirst, , BgPjt.ProjectName)
    End If
    BgPjt.IndexTrVw = lvNode.Index
    
End Sub


