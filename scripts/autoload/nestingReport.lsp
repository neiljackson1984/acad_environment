; prompts to select an entity, then reports on the name and layer of the entity and
; of all the block references (if any) that the entity is nested within, proceeding from the deepest level of nesting to the block definition that is being actively edited.
(defun C:nestingReport ( /
    result
    entity
    nestingPath
    report
    )
    
    (setq result (nentsel))
    (setq entity (vlax-ename->vla-object (car result))) ; entity is the most deeply nested entity that the user could possibly have been construed to have clicked. (we do not expect entity to be a block reference.)
    ; nesting path is a list of blockReferences, starting with the block reference that points to the block definition that owns entity, then the 
    ; block reference that points to the block definition that owns the previous block reference, and so forth.
    ; as you might expect, if entity is owned by the blocki definition that is currently being edited, then nestingPath is an empty list (i.e. nestingPath is nil).
    ; that is the parent of the entity, and then, if entity was being displayed because there is a reference poiunting to that block definition, the parent block definition of that block reference is include, and so forth.
    (setq nestingPath 
        (mapcar 
            'vlax-ename->vla-object
            (nth 3 result)
        )
    )

    (setq report "")
    (mapcar 
        '(lambda (x)
            (setq report
                (strcat report
                    (if (= report "") "You clicked a" "which is contained in ")
                    (vla-get-objectName x)
                    " "
                    "("
                        (if (vlax-property-available-p x 'effectiveName)
                            (strcat "name: \"" (vla-get-effectiveName x)  "\", ")
                            (if (vlax-property-available-p x 'name)
                                (strcat "name: \"" (vla-get-name x)  "\", ")
                                ""
                            )
                        )
                        (strcat "layer: " (vla-get-layer x))
                    ")"
                    "\n"
                )
            )
        )
        (append
            (list entity)
            nestingPath
        )
    )
    (setq report 
        (strcat report 
            "which is contained in the active block definition."
            "\n"
        )
    )
    (princ "\n")
    (princ report)
    (princ)
)
(princ)