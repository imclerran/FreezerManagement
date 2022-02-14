Private Sub FindBoxesButton_Click()
    shelfLife.Update_Expired
    FindBoxes.Find_Boxes
End Sub

Private Sub ResetButton_Click()
    shelfLife.Update_Expired
    Call FindBoxes.Reset_Pallets
End Sub

Private Sub ShowExpiredButton_Click()
    shelfLife.Update_Expired
End Sub
