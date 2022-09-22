;;Defines the command mtext_edit_externally, which opens the selected mtext object 
;; in notepad, regardless of the current setting of mtexted
(defun C:mtext_edit_externally ( / 
    initialMtexted
    )
    ; (vlax-for view (vla-get-Views (vla-get-ActiveDocument (vlax-get-acad-object)))
        ; (vlax-dump-object view)
    ; )
    
    ; (setq newView 
        ; (vla-Add 
            ; (vla-get-Views (vla-get-ActiveDocument (vlax-get-acad-object)))
            ; (GUID )
            ; ; (strcat "made while on layout " (vl-princ-to-string (vla-get-ObjectID (vla-get-ActiveLayout (vla-get-ActiveDocument (vlax-get-acad-object))))))
        ; )
    ; )
    ; (princ "\n")
    ; (princ "active layout: ") (princ  (vla-get-ObjectID (vla-get-ActiveLayout (vla-get-ActiveDocument (vlax-get-acad-object))))) (princ "\n")
    ; ; (princ "nome of new view: ") (princ  (vla-get-Name newView)) (princ "\n")
    ; (princ "LayoutID of view: ") (princ  (vla-get-LayoutID newView)) (princ "\n")
    (setq initialMtexted (getvar "MTEXTED"))
    (setvar "MTEXTED" "notepad")
    (command "_MTEDIT")
    (setvar "MTEXTED" initialMtexted)
)

(defun C:mtee ( / ) (C:mtext_edit_externally))

