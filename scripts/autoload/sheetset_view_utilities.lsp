;; makes a new layout view in the current layout, assigns a newly-generated guid as a name, and prompts to associate the newly created layout view with a viewport.
(defun C:viewAdd ( / 
    GUID
    uniqueName
    mySelectionSet
    initiallySelectedSelectionSet
    initialVariableValues
    initialErrorHandler
    cleanup
    pviewport
    i
    selectedPviewports
    entity
    sheetView
    result
    e
    weHaveFinishedPerformingTheFirstViewportAssociation
    existingAssociatedSheetview
    selectedTypes
    LM:Unique 
    implode
    origin
    ss
    )


    (progn ;error handling setup
        (setq initialVariableValues
            (mapcar
                '(lambda (variableName) (cons variableName (getVar variableName)))
                (list
                    "cmdecho"
                    "pickfirst"
                )
            )
        )
        (setq initialErrorHandler *error*)
        (defun cleanup ( )
            (mapcar '(lambda (x) (setvar (car x) (cdr x)) ) initialVariableValues)
            (setq *error* initialErrorHandler)
            (*pop-error-mode*) 
            (princ)
        )
        (defun *error* ( errorMessage )  
            (princ (strcat "we have encountered an exception: " errorMessage "\n"))
            (cleanup)
        )
        (*push-error-using-stack*) ; declares that the error handler should have access to local variables inside the function
        ; (princ "initialVariableValues: ")(princ initialVariableValues)(princ "\n")
   )
   ;======
    
    (progn ; private functions
        ;; thanks to https://www.theswamp.org/index.php?topic=41820.0
        (defun GUID  (/ tl g)
          (if (setq tl (vlax-get-or-create-object "Scriptlet.TypeLib"))
            (progn (setq g (vlax-get tl 'Guid)) (vlax-release-object tl) (substr g 2 36)))
        )
 
        ;; Unique  -  Lee Mac
        ;; Returns a list with duplicate elements removed.
        (defun LM:Unique ( l )
            (if l (cons (car l) (LM:Unique (vl-remove (car l) (cdr l)))))
        )
        ;==============
        
        
        (defun implode (inputList delimiter / 
                LM:lst->str
                returnValue
            )
            
            ;; List to String  -  Lee Mac
            ;; Concatenates each string in a supplied list, separated by a given delimiter
            ;; lst - [lst] List of strings to concatenate
            ;; del - [str] Delimiter string to separate each item
            (defun LM:lst->str ( lst del )
                (if (cdr lst)
                    (strcat (car lst) del (LM:lst->str (cdr lst) del))
                    (car lst)
                )
            )
            
            ;;LEE mac's LM:lst->str function is pathological in the case where the input list is an empty list (i.e. nil)
            ;; in that case, the function returns nil, and not a string.
            ;; that is why I have wrapped lee mac's version in my own version (implode) that corrects the error.
            (setq returnValue (LM:lst->str inputList delimiter))
            (setq returnValue 
                (if returnValue returnValue "")
            )
            returnValue
        )
    )
    ;======
    
    (setq initiallySelectedSelectionSet (cadr (ssgetfirst)))
    (setvar "cmdecho" 0) ; suppress stdout from the commands  
    ; (setq uniqueName  (GUID))
    (setq uniqueName (strcat (rtos (getvar "CDATE") 2 6) "-" (GUID))) ; we prefix with a timestamp to ensure that alphabetic order and chronological order are the same, which can be somewhat useful in certain situations
    (command-s "-view" "Save" uniqueName)

    
    (setq result 
        (vl-catch-all-apply 'vla-Item 
            (list 
                (vla-get-Views (vla-get-ActiveDocument (vlax-get-acad-object)))
                uniqueName
            )
        )
    )
    
    (if (vl-catch-all-error-p  result)
        (progn
            (princ (strcat " failed to create layout view with exception: " (vl-catch-all-error-message result)  "\n"))
        )
        (progn
            (setq sheetView result)
        )
    )

    (if (not sheetView )
        (progn
            (princ "we have failed to create a new sheet view. \n")
            (quit)
        )
    )
    (princ (strcat "We have created a layout view, whose name is \"" uniqueName "\"" "\n"))
    
    ;;keep on prompting for selection of objects until the user succeeds in selecting a viewport (or hits Escape to kill this whole Lisp execution)
    (if (and (/= (getvar "pickfirst") 0) initiallySelectedSelectionSet (> (sslength initiallySelectedSelectionSet) 0))
        (progn
            ; in the case we will attempt to findthe viewport in the users existing selection (which will be retrieved by (ssget) below
            ; we need to re-select the items in case the above act of creating the sheet view caused the selection set to be cleared.\
            (sssetfirst nil initiallySelectedSelectionSet)
        )
        (progn
            ;; in this case, we will prompt the user to pick a viewport
        )
    )
    
    (setq mySelectionSet nil)
    (setq weAreFinishedWithSelection nil)
    (while (not weAreFinishedWithSelection)
        (princ
            (strcat 
                "Select a viewport to be associated with the newly-created layout view " "\n"
                "(or press enter not to associate a viewport, " "\n"
                 "or select multiple viewports and we will ensure that each " 
                 "is associated with a layout view (creating new layout views as needed)):" " \n"
            )
        )
        (if (= (getvar "pickfirst") 0)
            ; deselect everything that is currently selected, so that when (ssget) (which is snesitive to pickfirst)
            ; prompts the user to select objects, the user willnot see the existing selected objects highlighted and 
            ; mistakenely beleived thet they  will be included in the selection set.
            (sssetfirst nil (ssadd))
        )
        (setq mySelectionSet (ssget)    )
        (if (or (not mySelectionSet ) (< (sslength mySelectionSet) 1))
            (progn
                ;in this case the user did not select anything, indicating that he does not want
                ; to associate the newly-created layout view with a viewport.
                (princ "you have selected nothing, so we will not associate the newly-created layout view with a viewport.\n")
                (setq selectedPviewports nil)
                (setq weAreFinishedWithSelection T)
            )
            (progn
                ;in this case, the user selected something, let's find all the viewports in his selection
                (setq i 0)
                ; (setq countOfSelectedPviewports 0)
                ; (setq pviewport nil)
                (setq selectedPviewports nil)
                (setq selectedTypes nil)
                (while (< i (sslength mySelectionSet))
                    ; (princ (strcat "now checking entity " (itoa i) "\n"))
                    (setq entity (vlax-ename->vla-object (ssname mySelectionSet i)))
                    (setq selectedTypes
                        (append
                            selectedTypes
                            (list (vla-get-ObjectName entity))
                        )                       
                    )
                    
                    (if (= (vla-get-ObjectName entity) "AcDbViewport")
                        (progn
                            (setq selectedPviewports 
                                (append
                                    selectedPviewports
                                    (list entity)
                                )
                            )                           
                        )
                    )
                    (setq i (+ i 1))
                )
                
                (setq selectedTypes
                    (acad_strlsort 
                        (LM:Unique selectedTypes)
                    )
                )
                
                (if selectedPviewports
                    (progn
                        ; in this case, the user has selected at least one viewport, which means we 
                        ; are finsihed with the veiwport selection process.
                        (if (= (sslength mySelectionSet) (length selectedPviewports))
                            (progn
                                (princ (strcat "You have selected " (itoa (length selectedPviewports) ) " " (if (> (length selectedPviewports) 1) "viewports" "viewport")  ".  Thank you.\n"))
                            )
                            (progn
                                (princ (strcat "You have selected " (itoa (sslength mySelectionSet)) " entities, of which " (itoa (length selectedPviewports) ) (if (> (length selectedPviewports) 1) " are viewports" "is a viewport") ".  Thank you.\n"))
                            )
                        )
                        
                        
                        (setq weAreFinishedWithSelection T)
                    )
                    (progn
                        (princ (strcat "You have selected " (itoa (sslength mySelectionSet)) " entities " "(types: " (implode selectedTypes ", ") ")" ", but none of them are viewports.  Please try again.\n"))
                        (setq weAreFinishedWithSelection nil)
                    )
                )  
            )
        ) 
    )
    ;=====
    

    (setq weHaveFinishedPerformingTheFirstViewportAssociation nil)
    (setq selectedPviewportsSelectionSet (ssadd))
    (foreach pviewport selectedPviewports
        (princ (strcat "now processing viewport " (vla-get-Handle pviewport) ": "))
        (setq selectedPviewportsSelectionSet (ssadd (handent (vla-get-Handle pviewport)) selectedPviewportsSelectionSet))
        (setq existingAssociatedSheetview nil) 
        (setq result 
            (vl-catch-all-apply 'vla-get-SheetView 
                (list 
                    pviewport
                )
            )
        )
        (if (vl-catch-all-error-p  result)
            (progn
                (princ 
                    (strcat
                        "The viewport appears to have an existing sheet view association, but retrieving the sheet view failed with an exception: " (vl-catch-all-error-message result)  "\n"
                        "Therefore, we will re-assign a sheet view to this viewport (creating a new sheet view if needed)." "\n"
                    )
                )
                (setq existingAssociatedSheetview nil)
            )
            (progn
                (setq existingAssociatedSheetview result)
            )
        )        
        
        
        
        
        
        ; (princ "checkpoint\n")
        (if 
            (and
                existingAssociatedSheetview
                (not (vlax-erased-p existingAssociatedSheetview))
            )
            (progn 
                ;in this case the viewport is already associated with a sheet view, so we will skip it.
                ; (princ "checkpoint2\n")
                (princ (strcat "the viewport " (vla-get-Handle pviewport) " is already associated with a sheet view (namely \""  (vla-get-Name existingAssociatedSheetview) "\"), so we will not modify it." "\n"))
            )
            (progn     
                ; (princ "checkpoint3\n")
                ;in this case, the viewport is not already associated with a sheet view, so we will associate it with one (creating one if needed)
                (if weHaveFinishedPerformingTheFirstViewportAssociation
                    (progn
                        ;in this case, we need to create a new sheetView:
                        (setq uniqueName (strcat (rtos (getvar "CDATE") 2 6) "-" (GUID))) ; we prefix with a timestamp to ensure that alphabetic order and chronological order are the same, which can be somewhat useful in certain situations
                        (command-s "-view" "Save" uniqueName)
                        (setq result 
                            (vl-catch-all-apply 'vla-Item 
                                (list 
                                    (vla-get-Views (vla-get-ActiveDocument (vlax-get-acad-object)))
                                    uniqueName
                                )
                            )
                        )
                        (if (vl-catch-all-error-p  result)
                            (progn
                                (princ (strcat " failed to create layout view with exception: " (vl-catch-all-error-message result)  "\n"))
                                (setq sheetView nil)
                            )
                            (progn
                                (setq sheetView result)
                            )
                        )
                    )
                )
                (if sheetView
                    (progn
                        (vla-put-SheetView pviewport sheetView)
                        (vla-put-HasVpAssociation sheetView :vlax-true)
                        (princ (strcat "we have succesfully created a new layout view " "(named \"" (vla-get-Name sheetView) "\")" " and associated the layout view with viewport " (vla-get-Handle pviewport) ".\n"))
                        
                        
                        ; jiggle the viewport to cause autocad to update the view's cropbox accordingly
                        ; (vla-Update pviewport)
                        
                        ;;DOESN't CAUSE THE VIEW TO CHANGE
                        ; ; (setq origin (vlax-3d-point 0 0 0 ))
                        ; ; (vla-Move pviewport 
                            ; ; origin
                            ; ; (vlax-3d-point 1 0 0)
                        ; ; )
                        ; ; (vla-Update pviewport)
                        ; ; (command-s "regen")
                        ; ; (vla-Move pviewport 
                            ; ; origin
                            ; ; (vlax-3d-point -1 0 0)
                        ; ; )
                        ; ; (command-s "regen")
                        ; ; (vla-Update pviewport)
                        
                        ;TOO SLOW to do the jiggling one by one.  I will jiggle all at once outseide the loop.
                        ; (setq ss (ssadd (handent (vla-get-Handle pviewport)) (ssadd)))
                        ; ; (princ (strcat "ss contains " (itoa (sslength ss)) " entities.\n"))
                        ; (setvar "pickfirst" 0)
                        ; (command-s "_MOVE" ss "" (list 0 0 0) (list 1 0 0))
                        ; (command-s "_MOVE" ss "" (list 0 0 0) (list -1 0 0))
                        ; (setq ss nil)

                        (setq weHaveFinishedPerformingTheFirstViewportAssociation T)
                    )
                )
            )
        )
    )
    ;; jiggle all selected viewports to make sure that AutoCAD updates the views accordingly.
    ;; TO DO: handle the case where the viewports are on locked layers gracefully.

    (setvar "pickfirst" 0)
    (command-s "_MOVE" selectedPviewportsSelectionSet "" (list 0 0 0) (list 1 0 0))
    (command-s "_MOVE" selectedPviewportsSelectionSet "" (list 0 0 0) (list -1 0 0))
    (setq selectedPviewportsSelectionSet nil)
    
    (cleanup)

)























; prompts to select a pviewport, then prompts for the name of a sheet view to set as the value of that pviewport's SheetView property.
(defun C:pviewportAssociateWithSheetView ( /
    pViewport
    pickedEntity
    pickedPoint
    viewportPickingPrompt
    nameOfSheetView
    sheetView
    activeDocument 
    GUID
    )
    
    ;; thanks to https://www.theswamp.org/index.php?topic=41820.0
    (defun GUID  (/ tl g)
      (if (setq tl (vlax-get-or-create-object "Scriptlet.TypeLib"))
        (progn (setq g (vlax-get tl 'Guid)) (vlax-release-object tl) (substr g 2 36)))
    )

    (setq activeDocument (vla-get-ActiveDocument (vlax-get-acad-object)))
    (setq viewportPickingPrompt "Please select a pViewport.")
    ; (setq sheetViewPickingPrompt "Enter the name of a sheet view (or leave blank to create a new sheet view)")
    (setq sheetViewPickingPrompt "Enter the name of a sheet view")
    (princ "\n")
   
    (setq pViewport (vlax-ename->vla-object (car (entsel viewportPickingPrompt))))
    (princ "\n")
    ;;TO-DO validate the selected entity to make sure it is a viewport

    (setq nameOfSheetView 
        (vla-GetString
            (vla-get-Utility activeDocument)
            1 ;; hasSpaces. 1 means that the input can have spaces
               sheetViewPickingPrompt
        )
    )
    

    
    ;;To-Do: validate the sheetView name input that we just retrieved
    ;;To-Do implement creating new sheet view if the nameOfSheetView is empty.
    (setq sheetView 
        (vla-item
            (vla-get-Views activeDocument)
            nameOfSheetView
        )
    )
    
    
    
    (vla-put-SheetView pViewport sheetView)
    (vla-put-HasVpAssociation sheetView :vlax-true)
    
    ; (vlax-dump-object sheetView)
    ; (vlax-dump-object pViewport)
    
    
    (princ)
)
(princ)
    ; (vlax-for view 
        ; (vla-get-Views (vla-get-ActiveDocument (vlax-get-acad-object)))
        ; (vlax-dump-object view)
    ; )
    
    
    ; (vlax-for view 
        ; (vla-get-Views (vla-get-ActiveDocument (vlax-get-acad-object)))
        ; (princ (vla-get-Name view))(princ "\n")
    ; )


    ; (vlax-dump-object
        ; (vla-get-SheetView
            ; (vlax-ename->vla-object (car (entsel)))
        ; )
    ; )
        ; (vlax-dump-object  (vlax-ename->vla-object (car (entsel))))
    ; ; 397d9adf58dd4ff4af2e55609ad501f1
    
; (vlax-dump-object (vla-item (vla-get-Views (vla-get-ActiveDocument (vlax-get-acad-object))) "397d9adf58dd4ff4af2e55609ad501f1"))
