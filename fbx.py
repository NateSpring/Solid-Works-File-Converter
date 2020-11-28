from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys
import pathlib
import codecs
import os, os.path
import time


# test file path
# \\filestore\\network storage\\Data\\ITD\\Team-ITD-Super-Site\\Images\\Product Renderings\\To Be Converted\ITD4096.SLDPRT

export_path = r'C:\\Users\\andrew foster.ITD\\Documents\\SOLIDWORKS Visualize Content\\Models'
dir_to_rer = r'\\filestore\\network storage\\Data\\ITD\\Team-ITD-Super-Site\\Images\\Product Renderings\\To Be Converted'
render_que = r"\\filestore\network storage\Data\ITD\Team-ITD-Super-Site\Images\Product Renderings\Render Queue"

sucessful_export = 0

for file in os.listdir(dir_to_rer):
    if file.lower().endswith(('.sldasm', '.sldprt')):
        file_count = len(os.listdir(dir_to_rer))
    #Handle File Name
        full_path = os.path.join(dir_to_rer, file)
        base = os.path.basename(full_path)
        file_stripped = os.path.splitext(base)[0]
        already_done = os.path.join(render_que, r'\%s\%s.fbx' % (file_stripped, file_stripped))
    #OPEN SolidWERX
        app = Application(backend="uia").start(r'C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS Visualize (2)\\SLDWORKSVisualize.exe')
        app = Desktop(backend="uia").window(title_re="SOLIDWORKS Visualize Standard")
        
        # if (os.path.exists(already_done)):
        #     print('%s already converted, moving on... ' % (file_stripped))
        #     continue
        # else:
        
        #FOLDER SETUP/ Export Path
        render_path = os.path.join(render_que, file_stripped)
        if (os.path.exists(render_path)):
            export_path = render_path + '\\'
            print('Directory Exists: %s' % (export_path))
            sucessful_export += 1
            print('Export Successful: File: %s of %s \n----------\n\n' % (sucessful_export, file_count))
            continue

        else:
            os.mkdir(render_path)
            export_path = render_path + '\\'
            print('Created Directory: %s' % (export_path))
        
            
        #NEW PROJECT
            app.menu_select('File->New')

            if (app.child_window(title="Save changes to Untitled Project.svpj before closing?").exists()):
                print('A pesky window appears... KILL')
                app.NoButton.click()
            
        #OPEN FILE    
            app.menu_select('File->Open')

            if (app.child_window(title="Save changes to Untitled Project.svpj before closing?").exists()):
                print('A pesky window appears... KILL')
                app.NoButton.click()
            app.OpenDialog.FileNameEdit.type_keys('%s' % (file))
            app.OpenDialog.Open3.click_input()
            app.child_window(title="Import Settings").wait('enabled', timeout=120)
            if (app.child_window(title="Import Settings").is_visible()):
                app.OK.click()
                print('Import Settings Accepted')

            if (app.child_window(title="Save changes to Untitled Project.svpj before closing?").exists()):
                print('A pesky window appears... KILL')
                app.NoButton.click()

        #Child Window Detection: No List Information
            if (app.child_window(title="This SOLIDWORKS file does not contain any display list information. Please open the file in SOLIDWORKS, save and try again.").exists()):
                app.OK.click()
                print('Found Error Window... Clicking OK!')
            
        ###Chapter 2: EXPORTING
            print('Export Window Engaged...')
            app.menu_select('File->Export->Export Project')
            app.child_window(title="Export").wait('enabled')
            app.AlllocationsSplitButton.click_input()

            app.window(title="Export").type_keys('^a{BACKSPACE}')
            app.window(title="Export").type_keys('%s' % (render_path), with_spaces=True )
            app.window(title="Export").type_keys('{ENTER}')


            if (app.child_window(title="Save changes to Untitled Project.svpj before closing?").exists()):
                print('A pesky window appears... KILL')
                app.NoButton.click()
            app.OpenDialog.FileNameEdit.type_keys('^a{BACKSPACE}')
            app.OpenDialog.FileNameEdit.set_text('%s' % (file_stripped))
            print('Export name set to: %s' % (file_stripped))
            app.OpenDialog.ComboBox2.select('Autodesk FBX Scene (*.fbx)')
            print('Export Type Selected: FBX')
            app.OpenDialog.Save.click()
        #Child Window Detection: Export Success
            if (app.child_window(title="Export of %s.fbx completed." % (file_stripped)).is_visible()):
                app.OK.click()
                sucessful_export += 1
                print('Export Successful: File: %s of %s \n----------\n\n' % (sucessful_export, file_count))
                app.menu_select('File->Close')
            if (app.child_window(title="Save changes to Untitled Project.svpj before closing?").exists()):
                print('A pesky window appears... KILL')
                app.NoButton.click()
    else:
        print('File is not SolidWorks... skipping')
        continue    

