import eel
import pandas as pd


file_path = r"C:\Tex Onsite\Tex_reff\shernof\thread-gage\Tex At Site Thread Gage Worksheet.xlsm"
df = pd.read_excel(file_path, sheet_name="Variables", engine="openpyxl" )
eel.init('GUI')
# items = None
# receive from JS
@eel.expose
def thread_gage(type):
    return type

# end of receive from JS

# to JS

# @eel.expose
# def export_items():
#     print(items)
#     return items

@eel.expose    
def get_model(type='metric'):
    # print(type)
    if type.lower() == "metric":
        series = df.iloc[16:58, 1] #metric
    else:
        series = df.iloc[89:119, 1] #imperial

    # global items
    items = [str(x).strip() for x in series.tolist() if pd.notna(x) and str(x).strip()]
    return items
get_model()

#end of to JS
if __name__ == "__main__":
    eel.start('index.html', size=(900, 600), position=(50,50))