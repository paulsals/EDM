# VBS Script Documentation

## Introduction

This document provides an overview and explanation of the VBS script that retrieves various board properties from an Expedition PCB application.

## Script Explanation

The script performs the following tasks:
1. **Get the Application object**: Retrieves the Expedition PCB application object.
2. **Get the active document**: Retrieves the active document from the application.
3. **License the document**: Validates the document license.
4. **Get the vias collection**: Retrieves the collection of vias in the document.
5. **Get the number of vias in collection**: Counts the number of vias.
6. **Get the number of Layers**: Counts the number of layers in the document.
7. **Outline coordinates**: Retrieves the board outline coordinates (MinX, MinY, MaxX, MaxY).
8. **Get the Name**: Retrieves the name of the document.
9. **Get Base Unit**: Retrieves the base unit of the document (IN or MM).
10. **Get Component total**: Retrieves the total count of components in the document.
11. **Get the ConnectionCountOption property**: Retrieves the connection count option property.
12. **Get the Nets property**: Retrieves the collection of nets in the document.
13. **Get the number of nets**: Counts the number of nets.
14. **Check for the presence of the "KANBAN" cell**: Checks if the "KANBAN" cell is present in the document.
15. **Get the smallest drill size of holes and count of non-plated holes**: Retrieves the smallest drill size and counts the non-plated holes.
16. **Calculate the new variable**: Calculates a new variable based on the total hole count, via count, and non-plated hole count.
18. **Gets the Number of Component with Ref Des of TP**: Retrieves the total count of components in the document with a Ref Des of TP.
19. **Gets the number of components with TP on Top of the board**: Retrieves the total count of components in the document with a Ref Des of TP on the top of the board.
20. **Gets the number of Components with TP on Bottom of the Board**: Retrieves the total count of components in the document with a Ref Des of TP on the bottom of the board. 
21. **Write the counts and board outline to a JSON file**: Writes the retrieved information to a JSON file.