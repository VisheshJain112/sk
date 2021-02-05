import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os","tkinter","numpy","tkmagicgrid","styleframe","pandas","matplotlib","csv","gingerit"]}
include_files = 'D:\desktop_app_scripts\Desktop_app\my_icon.ico'

# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(  name = "code",
        version = "1.0",
        description = "the python code",
        options = {"build_exe": build_exe_options},
        executables = [Executable("UserPanel.py", icon="myicon.ico", base=base)]
)



    for idx,row in self.total_df.iterrows():
        if get_category_value(idx)!=None:
            if back_category_value == get_category_value(idx):
                if remark_sheet[idx] == 0:
                    if df_attempt[idx] == 1:
                    if clb_sheet[idx]==1:

                        clb_sent.append(str(row['AN-Data dictionary']))
                        
                    else:

                        if row.Sequence!=row.Sequence:
                        pass
                        else:

                        if self.get_alphanumeric(str(row.Sequence)):
                            #print(row.Sequence)
                            data_dict[row.Sequence] = row['AN-Data dictionary']
                        else:
                            if "<" in row['A-Sentence'] :
                            
                            else:
                            data_dict[row.Sequence] = row['A-Sentence']

                    else:
                    pass
                else: