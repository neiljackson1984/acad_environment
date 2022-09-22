;;Defines the command mtext_edit_externally, which opens the selected mtext object 
;; in notepad, regardless of the current setting of mtexted
(defun C:uniquely_named_block ( / 
    initialMtexted
    GUID
    uniqueName
    mySelectionSet
    basePoint
    insertionPoint
    ; initiallySelectedSelectionSet
    LM:ssboundingbox
    boundingBox
    minPoint
    maxPoint
    midPoint
    initialVariableValues
    initialErrorHandler
    cleanup
    )

    (setq initialVariableValues
        (mapcar
            '(lambda (variableName) (cons variableName (getVar variableName)))
            (list
                "cmdecho"
                ; "pickfirst"
            )
        )
    )
    
    (setq initialErrorHandler *error*)
    
    (defun cleanup ( )
        (mapcar '(lambda (x) (setvar (car x) (cdr x)) ) initialVariableValues)
        (setq *error* initialErrorHandler)
        (princ)
    )
    
    (defun *error* ( errorMessage )  
        (princ (strcat "we have encountered an exception: " errorMessage "\n"))
        (cleanup)
    )
    (*push-error-using-stack*) ; declares that the error handler should have access to local variables inside the function
    
    (princ "initialVariableValues: ")(princ initialVariableValues)(princ "\n")
    
    ;; thanks to https://www.theswamp.org/index.php?topic=41820.0
    (defun GUID  (/ tl g)
      (if (setq tl (vlax-get-or-create-object "Scriptlet.TypeLib"))
        (progn (setq g (vlax-get tl 'Guid)) (vlax-release-object tl) (substr g 2 36)))
    )
    
    ;; Selection Set Bounding Box  -  Lee Mac
    ;; Returns a list of the lower-left and upper-right WCS coordinates of a
    ;; rectangular frame bounding all objects in a supplied selection set.
    ;; sel - [sel] Selection set for which to return bounding box

    (defun LM:ssboundingbox ( sel / idx llp ls1 ls2 obj urp )
        (repeat (setq idx (sslength sel))
            (setq obj (vlax-ename->vla-object (ssname sel (setq idx (1- idx)))))
            (if (and (vlax-method-applicable-p obj 'getboundingbox)
                     (not (vl-catch-all-error-p (vl-catch-all-apply 'vla-getboundingbox (list obj 'llp 'urp))))
                )
                (setq ls1 (mapcar 'min (vlax-safearray->list llp) (cond (ls1) ((vlax-safearray->list llp))))
                      ls2 (mapcar 'max (vlax-safearray->list urp) (cond (ls2) ((vlax-safearray->list urp))))
                )
            )
        )
        (if (and ls1 ls2) (list ls1 ls2))
    )
        
    (setq initialCmdecho (getvar "cmdecho"))
    (setvar "cmdecho" 0) ; suppress stdout from the commands
        
    ; (setq uniqueName (strcat (rtos (getvar "CDATE") 2 6) "-" (GUID))) ; we prefix with a timestamp to ensure that alphabetic order and chronological order are the same, which can be somewhat useful in certain situations
    (setq uniqueName 
        (vl-string-trim 
            " \t\n"
            (vl-string-subst 
                ""
                "-"                
                (strcase (GUID) T)
            )
        )
    ) ; we prefix with a timestamp to ensure that alphabetic order and chronological order are the same, which can be somewhat useful in certain situations

    ; (setq initiallySelectedSelectionSet (cadr (ssgetfirst)))
    
    (if (= (getvar "pickfirst") 0)
        ;deselect everything that is currently selected, so that when (ssget) (which is snesitive to pickfirst)
        ; prompts the user to select objects, the user willnot see the existing selected objects highlighted and 
        ; mistakenely beleived thet they  will be included in the selection set.
        (sssetfirst nil (ssadd))
    )
    ;;keep on prompting for selection of objects until the user succeeds in selecting something (or hits Escape to kill this whole Lisp execution)
    (while 
        (or 
            (not mySelectionSet) 
            (< (sslength mySelectionSet) 1)
        )
        (setq mySelectionSet (ssget)    )
    )
    
    (princ (strcat "you have selected " (itoa (if mySelectionSet (sslength mySelectionSet) 0)) " entities.\n"))
    ; (initget (+ 0 128) "default blab") ; allows the user to respond by pressing only Enter, in which case (get point ) will return an empty string.
    (setq basePoint (getpoint "Specify Base Point for Block (lower left corner of bounding box): "))
    ; (if (or (= (type basePoint) 'STR ) (not basepoint))
    (if (not basepoint)
        (progn
            (princ "you chose the default option: ")(princ basePoint)(princ "\n")
            (setq boundingBox (LM:ssboundingbox mySelectionSet))
            (setq minPoint (nth 0 boundingBox))
            (setq maxPoint (nth 1 boundingBox))
            (setq midPoint
                (mapcar '(lambda (a b) (/ (+ a b) 2)) minPoint maxPoint)
            )
            (setq basePoint minPoint)
        )
    )
        
 
    (setq insertionPoint basePoint)
    (princ "We will use ")(princ basePoint)(princ " as the base point of the block.\n")

    (command "-block" uniqueName basePoint mySelectionSet ""  )
    (command "-insert" uniqueName insertionPoint 1 1 0)
    (cleanup)
    (*pop-error-mode*) 
    (princ)
)

(defun C:ub ( / ) (C:uniquely_named_block))
(princ)
