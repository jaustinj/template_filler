import pandas as pd
import Tkinter
import tkFileDialog
import os

def file_choose():
    root = Tkinter.Tk()
    root.withdraw()
    filename = tkFileDialog.askopenfilename()
    return filename


class TrackExport(object):
    def __init__(self, csv_file):
        self.raw = self.clean_import(csv_file)

    def month(self, last_x_periods = False):
        month_df = self.report_pivot(self.raw,agg_code= 'M', last_x_periods = last_x_periods)
        return month_df

    def week(self, last_x_periods = False):
        week_df = self.report_pivot(self.raw,agg_code= 'W-SAT', last_x_periods = last_x_periods)
        return week_df

    def template_filler(self, last_x_periods = False):
        month_df = self.month(last_x_periods = last_x_periods)
        month_df['Month'] = 'Month'
        week_df = self.week(last_x_periods = last_x_periods)
        week_df['Week'] = 'Week'
        
        #Get values for months and weeks
        months = month_df.columns
        months = pd.DataFrame(months).transpose()
        weeks = week_df.columns
        weeks = pd.DataFrame(weeks).transpose()

        #change header names of month_df and week_df for appending
        month_df.columns = months.columns
        week_df.columns = weeks.columns

        #append them all together akwardly for report format
        awkward_template_df = months.append(weeks)
        awkward_template_df = awkward_template_df.append(month_df)
        awkward_template_df = awkward_template_df.append(week_df)

        col = awkward_template_df.columns.tolist()
        col = col[0:3] + col[-1:] + col[3:-1]

        awkward_template_df = awkward_template_df[col]

        return awkward_template_df
        
        

        

    def clean_import(self, csv_filepath):
        df = pd.read_csv(csv_filepath, header=None)
        
         #take timestamp, convert to datetime and set as index
        try:
            pd.to_datetime(df.iloc[3,0], format = '%Y-%m-%d %H:%M:%S')
            df.iloc[:,0] = pd.to_datetime(df.iloc[:,0], format = '%Y-%m-%d %H:%M:%S', errors = 'coerce') #format when downloaded
        except:
            df.iloc[:,0] = pd.to_datetime(df.iloc[:,0], format = '%m/%d/%y %H:%M', errors = 'coerce') #format if csv has been resaved in excel with a different name
        df = df.set_index(df.iloc[:,0])
        df.index.name = 'Date'
        df = df.iloc[:,1:]

        #Convert csv format to pandas multiindex format
        header_rows = range(0,3)

        for header_row in header_rows:
            valid_header_cell = ''
            valid_header = []

            for header_cell in df.iloc[header_row,:]:
                if not pd.isnull(header_cell):
                    try:
                        valid_header_cell = header_cell.split(": ")[1]
                    except:
                        valid_header_cell = header_cell
                valid_header.append(valid_header_cell)

            df.iloc[header_row,:] = valid_header

        multi_index_tuples = zip(df.iloc[0,:],df.iloc[1,:],df.iloc[2,:])
        df.columns = pd.MultiIndex.from_tuples(multi_index_tuples)
        df = df.iloc[3:,:]

        #set numbers as float instead of str
        df = df.iloc[:,:].apply(pd.to_numeric, errors='coerce')

        return df


    def report_pivot(self, df, agg_code = 'W-SAT', last_x_periods = False):
        #create platform pivot and a totals pivot(all platforms combined)
        df_platforms = self.easy_pivot(df,agg_code)
        df_totals = self.easy_pivot(df,agg_code)
        df_totals['Platform'] = 'Cross-Platform'
        df_totals = df_totals.groupby(['Brands', 'Platform', 'Measure']).sum().reset_index()

        #combine platform pivot and total pivot
        df_combined = pd.concat([df_platforms, df_totals], axis = 0)
        df_combined = df_combined.sort_values(by = ['Brands', 'Platform', 'Measure'])

        #allow for filtering to last 12 periods (For reporting perposes)
        if last_x_periods is not False:
            start_column = len(df_combined.columns)-(last_x_periods+1)
            end_column = len(df_combined.columns)-1
            df_combined = pd.concat([df_combined.iloc[:,0:3],df_combined.iloc[:, start_column:end_column]], axis = 1)
                                                                             
        return df_combined


    def easy_pivot(self, df, agg_code = 'W-SAT', columns_list=['Brands', 'Platform', 'Measure']):
        df = df.resample(agg_code).sum()
        df = df.unstack().reset_index()
        df.columns = ['Brands', 'Platform', 'Measure', 'Date', 'Value']
        df = df.pivot_table(values='Value',columns=columns_list,index='Date',aggfunc='sum')
        df = df.transpose().reset_index()
        return df

    



    
