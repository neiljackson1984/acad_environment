; prompts to select one or more block references.
; prompts for the name of a block definition.
; moidifies the block reference(s) so as to change which block definition the block reference references.
(defun C:change_reference ( / 
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
    desiredNameOfBlockDefinition
    desiredBlockDefinitionToReference
    blockReferencesToModify
    thisEname
    i
    blockReferencesThatDidNotNeedToBeModified
    blockReferencesThatWeModified
    result
    thisEntity
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
    
    (setq initialCmdecho (getvar "cmdecho"))
    (setvar "cmdecho" 0) ; suppress stdout from the commands
        
   
    (if (= (getvar "pickfirst") 0)
        ;deselect everything that is currently selected, so that when (ssget) (which is sensitive to pickfirst)
        ; prompts the user to select objects, the user will not see the existing selected objects highlighted and 
        ; mistakenly beleived thet they will be included in the selection set.
        (sssetfirst nil (ssadd))
    )
    
    
    ;;keep on prompting for selection of objects until the user succeeds in selecting something (or hits Escape to kill this whole Lisp execution)
    ; (while 
        ; (or 
            ; (not mySelectionSet) 
            ; (< (sslength mySelectionSet) 1)
        ; )
        ; (setq mySelectionSet (ssget)    )
    ; )
    
    (setq blockReferencesToModify (list ))
    ( while (< (length blockReferencesToModify) 1)
        
        (setq mySelectionSet (ssget)    )
        
        (setq i 0)
        (while (< i (sslength mySelectionSet))
            (setq thisEname (ssname mySelectionSet i))
            (setq thisEntity (vlax-ename->vla-object thisEname))
            (if (= (vla-get-ObjectName thisEntity) "AcDbBlockReference")
                (progn
                    (setq blockReferencesToModify 
                        (append
                            blockReferencesToModify
                            (list thisEntity)
                        )
                    )
                )
            )
            (setq i (+ 1 i))
        )
    )
    (princ (strcat "you have selected " (itoa (length blockReferencesToModify)) " block references.\n"))
    
    
    (setq desiredBlockDefinitionToReference nil)
    (while (not desiredBlockDefinitionToReference)
        (setq desiredNameOfBlockDefinition (getstring T "Specify the name of a block definition: "))
        (setq result 
            ( vl-catch-all-apply 
                'vla-Item
                (list 
                    (vla-get-Blocks (vla-get-ActiveDocument (vlax-get-acad-object))) 
                    desiredNameOfBlockDefinition
                )
            )
        )
        (if (vl-catch-all-error-p result)
            (progn
                (princ (strcat "the string '" desiredNameOfBlockDefinition "' does not seem to be resolvable to a block definition in the current document."))
            )
            (progn
                (setq desiredBlockDefinitionToReference result)
            )
        )
        
    )
    
    (setq blockReferencesThatDidNotNeedToBeModified (list))
    (setq blockReferencesThatWeModified (list))
    
    
    (foreach blockReferenceToModify blockReferencesToModify
        (if (= (vla-get-EffectiveName blockReferenceToModify) (vla-get-Name desiredBlockDefinitionToReference))
            (progn
                ; in this case, blockReferenceToModify already references the desired block definition, so there is no need to make a change
                (setq blockReferencesThatDidNotNeedToBeModified
                    (append
                        blockReferencesThatDidNotNeedToBeModified
                        (list blockReferenceToModify)
                    )
                )
            )
            (progn
                ; in this case, the block reference references some block definition other than the desired one, so we will
                ; modify the block reeference.
                (vla-put-Name blockReferenceToModify (vla-get-Name desiredBlockDefinitionToReference))
                (setq blockReferencesThatWeModified
                    (append
                        blockReferencesThatWeModified
                        (list blockReferenceToModify)
                    )
                )
            )
        )
    )
    (princ (strcat "processed " (itoa (length blockReferencesToModify)) " block references, and modified " (itoa (length blockReferencesThatWeModified)) " of them." ))
    (cleanup)
    (*pop-error-mode*) 
    (princ)
)

(princ)
