;; takes a selection of tables, and in each, clears all style overrides
(defun C:remove_table_style_overrides ( / 
    mySelectionSet
    initiallySelectedSelectionSet
    initialVariableValues
    initialErrorHandler
    cleanup
    table
    i
    selectedTables
    entity
    result
    e
    existingAssociatedSheetview
    selectedTypes
    LM:Unique 
    implode
    range
    plusPlus
    getContentCount
    rowIndex
    contentIndex
    content
    blockDefinition
    contentValue
    oldTextString
    newTextString
    LM:UnFormat
    enforcedSuffix
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
 
        ;; Unique  -  Lee Mac
        ;; Returns a list with duplicate elements removed.
        (defun LM:Unique ( l )
            (if l (cons (car l) (LM:Unique (vl-remove (car l) (cdr l)))))
        )
        ;==============
        
        
        ;;-------------------=={ UnFormat String }==------------------;;
        ;;                                                            ;;
        ;;  Returns a string with all MText formatting codes removed. ;;
        ;;------------------------------------------------------------;;
        ;;  Author: Lee Mac, Copyright © 2011 - www.lee-mac.com       ;;
        ;;------------------------------------------------------------;;
        ;;  Arguments:                                                ;;
        ;;  str - String to Process                                   ;;
        ;;  mtx - MText Flag (T if string is for use in MText)        ;;
        ;;------------------------------------------------------------;;
        ;;  Returns:  String with formatting codes removed            ;;
        ;;------------------------------------------------------------;;
        (defun LM:UnFormat ( str mtx / _replace rx )

            (defun _replace ( new old str )
                (vlax-put-property rx 'pattern old)
                (vlax-invoke rx 'replace str new)
            )
            (if (setq rx (vlax-get-or-create-object "VBScript.RegExp"))
                (progn
                    (setq str
                        (vl-catch-all-apply
                            (function
                                (lambda ( )
                                    (vlax-put-property rx 'global     actrue)
                                    (vlax-put-property rx 'multiline  actrue)
                                    (vlax-put-property rx 'ignorecase acfalse) 
                                    (foreach pair
                                       '(
                                            ("\032"    . "\\\\\\\\")
                                            (" "       . "\\\\P|\\n|\\t")
                                            ("$1"      . "\\\\(\\\\[ACcFfHLlOopQTW])|\\\\[ACcFfHLlOopQTW][^\\\\;]*;|\\\\[ACcFfHLlOopQTW]")
                                            ("$1$2/$3" . "([^\\\\])\\\\S([^;]*)[/#\\^]([^;]*);")
                                            ("$1$2"    . "\\\\(\\\\S)|[\\\\](})|}")
                                            ("$1"      . "[\\\\]({)|{")
                                        )
                                        (setq str (_replace (car pair) (cdr pair) str))
                                    )
                                    (if mtx
                                        (_replace "\\\\" "\032" (_replace "\\$1$2$3" "(\\\\[ACcFfHLlOoPpQSTW])|({)|(})" str))
                                        (_replace "\\"   "\032" str)
                                    )
                                )
                            )
                        )
                    )
                    (vlax-release-object rx)
                    (if (null (vl-catch-all-error-p str))
                        str
                    )
                )
            )
        )
        ;===
        
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
        
        ;; returns the number of content items in the specified table cell.
        (defun getContentCount ( theTable rowIndex columnIndex /
            maxContentIndexToCheck
            contentIndex
            numberOfContentItems
            )

            
            ;figure out how many content items this table cell contains (content items are indexed starting from zero.  By default, every cell has zero content items.  The stuff that the user inserts into a cell with the UI tends to end up in content item zero.
            (setq maxContentIndexToCheck 100) ;; a guard to prevent an endless loop, in case the function does not ever throw an exception.
            ;; step through contentIndices and call some function that needs a valid contentIndex until we hit an exception.
            (setq contentIndex 0)
            (while 
                (and 
                    (< contentIndex maxContentIndexToCheck)
                    (or
                        ; (not 
                            ; (vl-catch-all-error-p 
                                ; (setq result 
                                    ; (vl-catch-all-apply 
                                        ; ; 'vla-GetBlockTableRecordId2 (list myTable rowIndex columnIndex contentIndex) ;;returns zero without throwing exception on non-existent content items.
                                        ; 'vla-GetValue (list myTable rowIndex columnIndex contentIndex)
                                    ; )
                                ; )
                            ; )
                        ; ) 
                        ;; this is not too menaningful in testing for the existence of a content item because GetDataType2 says that the data type is acGeneral in the case where the content item does not exist.
                        ; (/=
                            ; (progn	
                                ; (vla-GetDataType2 myTable rowIndex columnIndex contentIndex 'dataType_out 'unitType_out)
                                ; dataType_out
                            ; )
                            ; acUnknownDataType
                        ; )
                        (/= (vlax-variant-type (vla-GetValue theTable rowIndex columnIndex contentIndex)) vlax-vbEmpty) ; vla-GetValue will return an empty variant in the case where the content contains a block reference.
                        (/= 
                            (vla-GetBlockTableRecordId2 theTable rowIndex columnIndex contentIndex)
                            0
                        )
                    )
                )
                (setq contentIndex (+ 1 contentIndex))
            )
            
            (setq numberOfContentItems contentIndex)
            
            numberOfContentItems
        )


        ;; generates the list (0 1 2 3 ... (rangeSize - 1))
        (defun range (rangeSize /
            i
            returnValue
            )
            (setq returnValue (list))
            (setq i 0)
            (while (< i rangeSize) 
                (setq returnValue (append returnValue (list i)))
                ;(setq i (+ 1 i))
                (plusPlus 'i)
            )
            returnValue
        )


        ;; this plusPlus function behaves like the postfix '++' operator in C++ and similar languages.  That is, it returns the initial value of x (the value Before incrementing).
        (defun plusPlus (x / initialValue returnValue)
            (if (= (type x) 'SYM)
                (progn
                    (setq returnValue (eval x))
                    (set x (+ 1 (eval x)))
                )
                (progn
                    (setq returnValue x) ;; in the case where x is not a symbol, we will return x.  This is just to be thorough - I only intend to use plusPlus x to operate on symbols.
                )
            )
            returnValue
        )
        
        
    )
    ;======
    
    (setq initiallySelectedSelectionSet (cadr (ssgetfirst)))
    (setvar "cmdecho" 0) ; suppress stdout from the commands  

    (setq mySelectionSet nil)
    (setq weAreFinishedWithSelection nil)
    (while (not weAreFinishedWithSelection)
        (princ
            (strcat 
                "Select one or more tables for which you want to remove all tableStyle ovrrides." "\n"
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
                (princ "you have selected nothing.\n")
                (setq selectedTables nil)
                (setq weAreFinishedWithSelection T)
            )
            (progn
                ;in this case, the user selected something, let's find all the tables in his selection
                (setq i 0)
                ; (setq countOfselectedTables 0)
                ; (setq pviewport nil)
                (setq selectedTables nil)
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
                    
                    (if (= (vla-get-ObjectName entity) "AcDbTable")
                        (progn
                            (setq selectedTables 
                                (append
                                    selectedTables
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
                
                (if selectedTables
                    (progn
                        ; in this case, the user has selected at least one table, which means we 
                        ; are finsihed with the  selection process.
                        (if (= (sslength mySelectionSet) (length selectedTables))
                            (progn
                                (princ (strcat "You have selected " (itoa (length selectedTables) ) " " (if (> (length selectedTables) 1) "tables" "table")  ".  Thank you.\n"))
                            )
                            (progn
                                (princ (strcat "You have selected " (itoa (sslength mySelectionSet)) " entities, of which " (itoa (length selectedTables) ) (if (> (length selectedTables) 1) " are tables" "is a table") ".  Thank you.\n"))
                            )
                        )
                        
                        
                        (setq weAreFinishedWithSelection T)
                    )
                    (progn
                        (princ (strcat "You have selected " (itoa (sslength mySelectionSet)) " entities " "(types: " (implode selectedTypes ", ") ")" ", but none of them are tables.  Please try again.\n"))
                        (setq weAreFinishedWithSelection nil)
                    )
                )  
            )
        ) 
    )
    ;=====
    
    (foreach table selectedTables

        
        ;; strip text formatting and ensure that ecah cell whose type is "Data" ends with "\n ", to aid in vertical spacing
        (progn 
            (foreach rowIndex (range (vla-get-rows table))
                (foreach columnIndex (range (vla-get-columns table))
                    (princ (strcat "cell style: " (vla-GetCellStyle table rowIndex columnIndex) "\n"))
                    
                    (foreach contentIndex (range (getContentCount table rowIndex columnIndex))
                        (setq content (vla-GetValue  table rowIndex columnIndex contentIndex))
                        (setq contentValue 
                            (vl-catch-all-apply 'vlax-variant-value (list content))
                        ) 
                        (princ "\t\t\t\t")(princ "content ")(princ contentIndex)(princ ": ")
                        (princ
                            (cdr 
                                (assoc 
                                    (vlax-variant-type content)
                                    variantType_enumValues
                                )
                            )
                        )
                        (princ " ")
                        (princ
                            (cdr 
                                (assoc 
                                    (progn	
                                        (vla-GetDataType2 table rowIndex columnIndex contentIndex 'dataType_out 'unitType_out)
                                        dataType_out
                                    )
                                    (list
                                        (cons  acBuffer          "acBuffer"            )
                                        (cons  acDate            "acDate"              )
                                        (cons  acDouble          "acDouble"            )
                                        (cons  acGeneral         "acGeneral"           )
                                        (cons  acLong            "acLong"              )
                                        (cons  acObjectId        "acObjectId"          )
                                        (cons  acPoint2d         "acPoint2d"           )
                                        (cons  acPoint3d         "acPoint3d"           )
                                        (cons  acResbuf          "acResbuf"            )
                                        (cons  acString          "acString"            )
                                        (cons  acUnknownDataType "acUnknownDataType"   )
                                    )
                                )
                            )
                        )
                        (princ " ")
                        (if (/= 0 (vla-GetBlockTableRecordId2 table rowIndex columnIndex contentIndex))
                            (progn
                                (setq blockDefinition (vla-ObjectIDToObject (vla-get-Document table) (vla-GetBlockTableRecordId2 table rowIndex columnIndex contentIndex)))
                                (princ (vla-get-ObjectName blockDefinition))(princ ": ")
                                (princ (vla-get-Name blockDefinition))
                                ;;(vlax-dump-object blockDefinition)
                            )
                            (progn	
                                ;; in this case we assume that the content is a text
                                (princ "content: ")(princ contentValue)(princ "\n")
                                (setq oldTextString (vl-catch-all-apply 'vla-GetTextString (list table rowIndex columnIndex contentIndex)))
                                (setq newTextString (LM:UnFormat oldTextString T))
                                ;; in certain special cases, we will append a suffix to newTextString
                                (if 
                                    (and
                                        (= (getContentCount table rowIndex columnIndex) 1) ;;there is only one content in this cell (and it's a text content)
                                        ; (= contentIndex (- (getContentCount table rowIndex columnIndex) 1)) ;;this is the last content in this cell
                                        ; (= (vla-GetCellStyle table rowIndex columnIndex) "Data") ;;this cell is of type "Data"  //this doesn't work as expected.  In my case, GetCellStlye is returing an empty string when I was expecting it to retrun "Data"
                                        ; (= (vla-GetCellStyle table rowIndex columnIndex) "_DATA") ;;this cell is of type "Data"  //this doesn't work as expected.  In my case, GetCellStlye is returing an empty string when I was expecting it to retrun "Data" //it looks like GetCellStyle returns "_DATA" to represent the bnuiilt-in "Data" style.
                                        ; // it looks like GetCellStlye returns an empty string if the cell stly is seet to "By Row/Column" in the UI
                                        
                                        ;;here is a hacky work-around that suits my immediate purpose, but ought to be made more robust in the future:
                                        (= (vla-GetRowType table rowIndex) acDataRow )
                                        
                                    )
                                    (progn 
                                        (princ "appending the spacing suffix.\n")
                                        (setq newTextString
                                            (strcat 
                                                (vl-string-trim "\t\n " newTextString) ;; trim newlines tabs and spaces from the beginning and end of the string
                                                "\\P " ;;append a newline and a space.
                                            )
                                        )
                                    )
                                    (progn 
                                        (princ "not appending the spacing suffix.\n")
                                        ; don't do anything in this case
                                    )
                                )
                                
                                
                                (princ "oldTextString: " )(princ oldTextString)(princ "\n")
                                (princ "newTextString: " )(princ newTextString)(princ "\n")
                                ;(princ "GetTextString: ")(princ (vl-catch-all-apply 'vla-GetTextString (list table rowIndex columnIndex contentIndex)))
                                
                                
                                
                                (vla-SetTextString table rowIndex columnIndex contentIndex 
                                    newTextString
                                )
                            )
                        )
                        
                        (princ "\n")
                        ;(vla-GetBlockTableRecordId2 table rowIndex columnIndex)
                    )
                    ;====
                    
                )
            )



        )
        ;=====
        

        (vla-ClearTableStyleOverrides table 0)  ;;the '0' specifies that we want to delete both table overrides and cell-overrides
        (vla-put-Height table 0.01)    


        
    )
 
    (setq mySelectionSet nil)
    
    (cleanup)
)

(princ)

; (vlax-dump-object  (vlax-ename->vla-object (car (entsel))))

; (C:remove_table_style_overrides)

