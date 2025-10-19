- create a new tab when draging and dropping an folder onto the tab bar (where no other tabs are present, if tabs are present where the folder is dropped, it should fall back to the copy/move routine).
- there should be a group edit option in the main context menu of the toolbar for editing group colors and names.
- add 'New Tab' option to the context menu whe right clicking in the empty space where no tabs are present on the toolbar which opens a new tab to 'My PC'.
fully implement the options dialog so that the 'Options' context menu option opens the dialog.
-- the option dialog should have tabs for different option grouppings for the following features:
1.) tab 1 should be 'Main Options':
1a.) checkbox: always re-open last session on crash or force closure of explorer.
1b.)  checkbox (with example of how it works): save group tab paths when the group is closed to their current path instead of always opening to the default set path for each tab. (this means, if the user opens group 1 which originally opens tab1 to c:\test\  and the user opens the 'testsubfolder1' and closes the group, if checked the next time the group is opened, tab1 should open to c:\test1\testsubfolder1\  instead of c:\test1\. otherwise if unchecked, it should always open to c:\tab1\ when the group is re-opened.)
2.) tab 2 should be 'group/island management' which contains a list widget which lists all the currently available groups, a 'New Group' button, 'Edit Group' button and 'Remove Group' button.
2a.)  list box: the group list should allow the user to double click an entry to open the group modification dialog (see specs for edit group button).
2b.)  button: New Group button should open the default 'Create new group' dialog that is already present in the codebase.
2c.)  button: the edit group button will open a secondary add / edit dialog.
2d.) button: a button to add a new group will open a secondary add / edit dialog.  
2e.) button: a button to deleted the selected group from listbox.
3.) secondary add / edit dialog:
3a.)  listbox: contains all the tabs / paths to open tabs to when opening a group. New groups should have a default single entry set to 'c:\'.
3b.) button: 'Edit Path' will open the 'open folder' dialog which allows the user to navigate to the folder path they want and replace the currently selected path from the listbox. (P.S. this button should be disabled if no entry is selected.) 
3c.) button: 'Add Path' will add a new tab / path to the group, by opening a 'open folder' dialog which allows the user to navigate to the folder path they want to add.
3d.) button: 'Remove Path' will remove the selected path from the group. (P.S. this button should be disabled if no entry is selected.)
3e.) textbox: prefilled textbox for the group name if editing a group, if new default to 'New Group #'.
3f.) color picking widget: pre-picked color, color picker widget which defaults to the groups current color if editing group else default to a random color if new.
4.) buttons that should always be visible at the bottom of the options dialog despite the current tab:
4a.) buttons: Save button to apply changes.
4b.) button: Cancel button to discard changes.

-----------------

- fix the ability to add highlighting to files and folders.
- fix ability to change file and folder name font color based on associated tags.
- create the tag color system for file and folder name color system.
- make it so all views in the file explorer and save/open dialogs are hooked properly to show the color and highlighting of files and folders.
- implement session history which is saved periodically so if explorer crashes, is forced closed or any other issue happens which causes explorer to close, it will open up to the same tabs that where present before the forced closing of explorer.
- fix the filename and foldername color assignment code AND tag assignment code.  it currently does not detect the files and folders that are highlighted and thinks nothing is selected. this may involve implementing more file explorer hooks or com hooks.
- add the options menu to the right click context menu when right clicking on a tab and tab island.

------------

- tabs only move to the second row when minimizing and then un-minimizing the file explorer window. when doing this, it causes the tabs to get much thicker but moves them to the next row when they dont fit / arent visible on the toolbar due to the window size.
- fix the toolbar and tab resizing both horizontal and vertical.  disable the ability to make tabs thicker, the toolbar should be able to get taller unless tabs are off screen and need to be moved to a different row to be visible.
- find and fix multiple memory leaks.  eventually after enough time, icons disapear and the file explorer lags and eventually forces explorer to crash.
-------------
- fix the island group outline so it also encapsolates the island indicator line that is before the first tab in the island.
- fix the way the outline for island groupings work when it is broken into more than 1 row: instead of making the outline span 2 rows, the outline should only ever be 1 row tall so when a group is split between two rows, the outline should be split into 2 seperate outlines, 1 on the groupping in the top row and one for the groupping in the bottom row.  otherwise the way it currently works is the outline is as tall as the amount of rows it takes up which makes different islands look like they are within the split island.
- fix the options dialog so the tabs have labels on them.
- fix the group/island options tab so the list of islands/groups is an actual list box and that it actually displays all the saved group names.
- fix the edit groups dialog so that the path list box is an actually list box component and displays all the different tab paths for the specific island/group.

-------------------------

- fully remove the 'split tab comparison' code.
- fully remove the 'filename color' code as well.
- create a comprehensive plan for investigating, hooking and implementing the ability to modify the left-tree view pane and the main file view pane gui components of windows 10 file explorer.exe. this means creating the correct hooking code to be able to fully change things like the font used for folder and file names, the color of font within the two panes, the ability to implement highlighting of files, the ability to fully control both panes.
-----------------------------

the current pane hooking system and filename color change and hook system are completely broken and useless.
we need to start fresh with both ideas, starting with how and what we need to hook to properly get things to work such as recognizing when an item in the panes is selected.  currently things like the file name color changer doesnt detect anything as being selected in the main view pane when it is clearly selected.  the splitting of the tabs is horribly broken as well.
create a hook and start implementing the ability to modify the left-tree view pane and the main file view pane gui components of windows 10 file explorer.exe. this means creating the correct hooking code to be able to fully change things like the font used for folder and file names, the color of font within the two panes, the ability to implement highlighting of files, the ability to fully control both panes. 

the goal is to implement it into the tab file explorer extension to add things such as actual filename highlighting in the panes, filename text colors and much more.

use all of the information from the document in the root of the repository, it contains instructions for the entire process and all the information you need to figure this whole thing out: 
Hooking Windows Explorer’s Tree and List View Panes for Custom Drawing.docx
--------------

thoroughly analyze this document 'Hooking Windows Explorer’s Tree and List View Panes for Custom Drawing.docx' and then figure out how we can correctly and fully fix the filename colored highlighting and filename color changing so it works when items are selected and its visible in both file explorer panes for the specific file that has those things set on them.