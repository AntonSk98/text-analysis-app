from interface.appjar import gui
from service.serviceAnalysis import *


def press(button):
    if button != 'Quit':
        if not are_entries_empty(app.getAllEntries(), app):
            merge_csv_files(app.getEntry('directory_to_analyze'))
            get_dictionary_of_words(app.getEntry('file_to_analyze'))
            get_all_data_from_merged_document()
            if button == 'Analyze by texts':
                run_analyzer('texts')
                save_result_file(app.getEntry('directory_for_results'), 'text_analysis')
                app.infoBox(title='Analyze by texts', message='Analyze by texts is completed!')
            elif button == 'Analyze by topics':
                run_analyzer('posts')
                save_result_file(app.getEntry('directory_for_results'), 'post_analysis')
                app.infoBox(title='Analyze by topics', message='Analyze by topics is completed!')
            else:
                run_analyzer('texts & posts')
                save_result_file(app.getEntry('directory_for_results'), 'text_and_post_analysis')
                app.infoBox(title='Analyze by texts & topics', message='Analyze by texts & topics is completed!')
    else:
        app.stop()


app = gui("Antoszinsky's A.T.A.", useTtk=True, showIcon=False)
app.setSize(600, 200)
app.setTtkTheme(app.getTtkThemes()[1])
app.setResizable(canResize=False)

app.setPadding('40', '0')
app.addEmptyLabel('empty_label1')
app.addLabel("Choose a folder with texts to analyse (files must be in csv format).")
app.addDirectoryEntry('directory_to_analyze')
app.addLabel("Choose a dictionary (it must be an excel file)")
app.addFileEntry('file_to_analyze')
app.addLabel("Choose a folder to save the result")
app.addDirectoryEntry('directory_for_results')
app.addEmptyLabel('empty_label2')
app.addButtons(["Analyze by texts", "Analyze by topics", "Analyze by texts & topics", "Quit"], press)
app.addEmptyLabel('')
app.go()
input()
