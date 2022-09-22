
(defun renameBlocksAndLayers ( oldName newName /
    layerToBeRenamed
    blockDefinitionToBeRenamed
    existingLayerHavingNewname
    existingBlockDefinitionHavingNewname
    x
    blockDefinition
    entity
    )
    
    
    (setq layerToBeRenamed  
        (if
            (vl-catch-all-error-p 
                (setq x
                    (vl-catch-all-apply 'vla-Item
                        (list (vla-get-Layers (vla-get-ActiveDocument (vlax-get-acad-object))) oldName)
                    )
                )
            )
            nil
            x
        )
    )
    
    (setq blockDefinitionToBeRenamed  
        (if
            (vl-catch-all-error-p 
                (setq x
                    (vl-catch-all-apply 'vla-Item
                        (list (vla-get-Blocks (vla-get-ActiveDocument (vlax-get-acad-object))) oldName)
                    )
                )
            )
            nil
            x
        )
    )
    
        
    (setq existingLayerHavingNewname  
        (if
            (vl-catch-all-error-p 
                (setq x
                    (vl-catch-all-apply 'vla-Item
                        (list (vla-get-Layers (vla-get-ActiveDocument (vlax-get-acad-object))) newName)
                    )
                )
            )
            nil
            x
        )
    )
    
    (setq existingBlockDefinitionHavingNewname  
        (if
            (vl-catch-all-error-p 
                (setq x
                    (vl-catch-all-apply 'vla-Item
                        (list (vla-get-Blocks (vla-get-ActiveDocument (vlax-get-acad-object))) newName)
                    )
                )
            )
            nil
            x
        )
    )
    
    
   
    
    
    
    (if layerToBeRenamed
        (progn
            (if existingLayerHavingNewname
                (progn
                    ; merge layerToBeRenamed into existingLayerHavingNewname
                    (vlax-for blockDefinition (vla-get-Blocks (vla-get-ActiveDocument (vlax-get-acad-object)))
                        (vlax-for entity blockDefinition
                            (if (= (vla-get-Layer entity) oldName)
                                (progn
                                    (vla-put-Layer entity newName)
                                )
                            )
                        )
                    )
                    (vla-Delete layerToBeRenamed)
                )
                (progn
                    (vla-put-Name layerToBeRenamed newName )
                )
            )
        )        
    )
    
    (if (and blockDefinitionToBeRenamed (not existingBlockDefinitionHavingNewname))
        (progn
            (vla-put-Name blockDefinitionToBeRenamed newName )
        )        
    )
)
