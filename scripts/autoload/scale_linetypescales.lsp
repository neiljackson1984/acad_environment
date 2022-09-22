;;Modifies the "LinetypeScale" property of the selected objects by multiplying by a scale factor,
; but only touches entities whose resolved linetype (that is, linetype after resolving the "ByLayer" linetype)
; is something other than continuous or ByBlock.
; does not descend into block references. does not touch block references (because autocad completely ignores, and
; does not even allow the user to modify in the ui, the LinetypeScale property of a block reference.
(defun C:scale_linetypescales ( / 
    mySelectionSet
    scaleFactor
    objectToString
    initialVariableValues
    initialErrorHandler
    cleanup
    i
    resolvedLinetype
    counter
    initialLinetypeScale
    newLinetypeScale
    entity
    getOwner
    )

    (setq initialVariableValues
        (mapcar
            '(lambda (variableName) (cons variableName (getVar variableName)))
            (list
                ; "cmdecho"
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
    
    ; (princ "initialVariableValues: ")(princ initialVariableValues)(princ "\n")

    (defun objectToString (object /
            nameOfObject
            result
        )

        ; (setq result (vl-catch-all-apply 'vla-get-Name object))
        ; (if (vl-catch-all-error-p  result)
            ; (progn
                ; (setq nameOfObject nil)
            ; )
            ; (progn
                ; (setq nameOfObject result)
            ; )
        ; )
        (setq nameOfObject nil)
        (if
           (vlax-property-available-p object 'Name)
           (progn
                (setq nameOfObject 
                    (vl-princ-to-string
                        (vlax-get-property object 'Name)
                    )
                )
            )
        )
        
       
        (strcat 
            (vla-get-ObjectName object) 
            "[" (vla-get-Handle object) "]" 
            (if nameOfObject (strcat "(" nameOfObject ")") "")
        )
    )

    (defun getOwner (entity)
        (vla-ObjectIDToObject (vla-get-Document entity) (vla-get-OwnerID entity))
    )

    (if (= (getvar "pickfirst") 0)
        ;deselect everything that is currently selected, so that when (ssget) (which is snesitive to pickfirst)
        ; prompts the user to select objects, the user will not see the existing selected objects highlighted and 
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
    (setq scaleFactor (getreal "Specify a scale factor: "))

    (princ (strcat "We will multiply the LinetypeScale property value of each of the selected entities (that need treatment) by " (vl-prin1-to-string scaleFactor) "\n"))

    (setq counter 0)
    (setq i 0)
    (while (< i (sslength mySelectionSet))
        (setq entity (vlax-ename->vla-object (ssname mySelectionSet i)))

        (setq resolvedLinetype (vla-get-Linetype entity))
        (cond
            ((= resolvedLinetype "ByBlock")
                (setq resolvedLinetype nil)
            )
            ((= resolvedLinetype "ByLayer")
                (if
                    ; We will attempt to resolve a "ByLayer" linetype only in case when the entity is in a "top-level" block definition (i.e. modelspace or one of the layouts)
                    ; or is not on layer 0. (because an entity that is on layer 0, has a linetype of bylayer, and is in a non-top-level block will inherit its linetype from a block reference)
                    (or
                        (= (vla-get-Handle (getOwner entity)) (vla-get-Handle (vla-get-ModelSpace (vla-get-Document entity))))
                        (= (vla-get-IsLayout (getOwner entity)) :vlax-true)
                        (/= (vla-get-Layer entity) "0")
                    )
                    (progn
                        (setq resolvedLinetype
                            (vla-get-Linetype
                                (vla-Item
                                    (vla-get-Layers (vla-get-Document entity))
                                    (vla-get-Layer entity)
                                )
                            )
                        )
                    )
                    (progn
                        (setq resolvedLinetype nil)
                    )
                )
            )
        )

        (if 
            (and
                resolvedLinetype
                (not (member resolvedLinetype (list "Continuous" "ByLayer" "ByBlock")))
                (not (= (vla-get-ObjectName entity) "AcDbBlockReference"))
            )
            (progn
                (setq initialLinetypeScale (vla-get-LinetypeScale entity)) 
                ; (princ (strcat "initialLinetypeScale of " (objectToString entity) " (whose resolved linetype is " resolvedLinetype ":  "  (vl-princ-to-string initialLinetypeScale) "\n"))
                (setq newLinetypeScale (* initialLinetypeScale scaleFactor))
                
                ; (princ (strcat "now attempting to change linetypescale of " (objectToString entity) " (whose resolved linetype is " resolvedLinetype " ) from "  (vl-princ-to-string initialLinetypeScale) " to "   (vl-princ-to-string newLinetypeScale)  "\n"))
                
                (setq result
                    (vl-catch-all-apply 'vla-put-LinetypeScale 
                        (list
                            entity
                            newLinetypeScale
                        )
                    )
                )
                (setq counter (+ counter 1))
                
                (if 
                    (and
                        (not (vl-catch-all-error-p result))
                        (= (vla-get-LinetypeScale entity) newLinetypeScale)
                    )
                    (progn
                        (princ (strcat "succeeded in changing linetypescale of " (objectToString entity) " (whose resolved linetype is " resolvedLinetype " ) from "  (vl-princ-to-string initialLinetypeScale) " to "   (vl-princ-to-string newLinetypeScale)  "\n"))
                    )
                    (progn
                        (princ (strcat "attempted but failed to change linetypescale of " (objectToString entity) " (whose resolved linetype is " resolvedLinetype " ) from "  (vl-princ-to-string initialLinetypeScale) " to "   (vl-princ-to-string newLinetypeScale) "\n"))
                        ()
                    )
                )
            )
        )
          
        (setq i (+ i 1))
    )
    
    (princ (strcat "modified the LinetypeScale of " (itoa counter) " entities." "\n"))
    
    
    
    (cleanup)
    (*pop-error-mode*) 
    (princ)
)

(princ)
