(defun C:removeSheetsetAssociation ( / 
    )
    (dictremove (namedobjdict) "AcSheetSetData")
    (princ "\nsheetset association has been removed.\n")
    (princ)
)
(princ)