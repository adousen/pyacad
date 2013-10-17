#!/usr/bin/env python
# -*- coding: utf-8 -*-
from pyacad.api import *
def addSupportPath (file_dir = '', PosNumber = 0):
    tmp = ''
    preferences = ThisDrawing.Application.Preferences
    currSupportPath = preferences.Files.SupportPath
    print(currSupportPath)
    preferences.Files.SupportPath = file_dir +";" + currSupportPath
    print(preferences.Files.SupportPath)
				
def menuLoad (fileName):
    try:
        ThisDrawing.Application.MenuGroups.Item(fileName)
    except Exception:    
        ThisDrawing.Application.MenuGroups.Load(fileName+".mnu")
        mnuGroup = ThisDrawing.Application.MenuGroups.Item(fileName)
        count = mnuGroup.Menus.Count
        for i in range(count):
            mnuGroup.Menus.Item(i).InsertInMenuBar(ThisDrawing.Application.MenuBar.Count + 1)    
    
