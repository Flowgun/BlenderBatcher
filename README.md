# BlenderBatcher

====================================================================

		BlenderBatcher.ahk	Version: 0.1

    an Autohotkey Script to run Python scripts on all Blend files in a given directory

		By: Nidhal Flowgun (BoubakerNidhal@hotmail.com)

====================================================================
	
	Description: This script allows to select one Python Script and a directory that has blend files
	and it will automatically run the selected Python script on all blend files in the selected directory
	and all of its subfolders.
	
	The script allows to resume the progress of any chosen Python script, instead of having to restart the operations
	on all the blend files. It does this by storing the name of the completed files in a TXT file that corresponds to
	the Python script. The TXT is saved where this autohotkey script is.
	
	The Python script is run on all the found blend files in the background. It uses the Blender installation that has its path
	selected. This Autohotkey script provides an option to run Blender without any of its addons to speed up the process.
	It does this by renaming the Addon folder so that Blender doesn't find it, and when the operation is complete or aborted,
	it renames it back to its original format.


====================================================================

