import os
import sys
import pandas as pd
import trackDash as td

reload(sys)
sys.setdefaultencoding('UTF8')

#Prompt User for filename
path_to_desktop = os.path.expanduser('~') + '/Desktop/'
file_name = path_to_desktop + td.get_filename()


#Import Track Dashboard Information
track_export = td.TrackExport(td.file_choose())


#Begin Writing Excel
writer = pd.ExcelWriter(file_name, engine='xlsxwriter') 
track_export.template_filler(13).to_excel(writer,sheet_name = 'Aggregations', startrow = 0, startcol = 0)
writer.save()
