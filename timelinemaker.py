import xlsxwriter
from datetime import timedelta, date
import pandas as pd

def daterange(date1, date2):
    for n in range(int ((date2 - date1).days)+1):
        yield date1 + timedelta(n)

                # mảng, hàng_event, cột_event, chiều_dài_event
def is_avaiable_event(arr, event_row_postion, event_col_postion, end_col_position):
    for i in range(event_col_postion, end_col_position+1):
        if not arr[event_row_postion][i] == 0:
            return False
    return True


def draw_event(arr, event_name, start_col_position, end_col_position, timeline_row_position):
    count = 1
    direction = -1
    row_postion = timeline_row_position + (count * direction)
    while True:
        if is_avaiable_event(arr, row_postion, start_col_position, end_col_position):
            break

        if not is_avaiable_event(arr, row_postion, start_col_position, end_col_position):
            direction *= -1
            row_postion = timeline_row_position + (count * direction)

        if not is_avaiable_event(arr, row_postion, start_col_position, end_col_position):
            direction *= -1
            count+=1
            row_postion = timeline_row_position + (count * direction)

    for i in range(start_col_position, end_col_position+1):
        arr[row_postion][i] = event_name

def get_format(value):

    range_color = [
        '#ffffff',
        '#ccffcc',
        '#ccffff',
        '#ccccff',
        '#ffccff',
        '#ffcccc',
        '#cc9900'
    ]

    # if type(value) is str:
    #     color_index = len(value) // 6 - 1
    # else:
    #     color_index = int(value) // 6 - 1
    
    if not int(value) is None:
        color_index = int(value)
    else:
        color_index = 0


    form = workbook.add_format({
    'bold':     True,
    'border':   6,
    'align':    'center',
    'valign':   'vcenter',
    'fg_color': range_color[color_index],
    'size': 20
    })
    return form

sheets_name = ['AOV', 'FF']
workbook = xlsxwriter.Workbook('Event timeline.xlsx')

for n in sheets_name:
    df = pd.read_excel("Events.xlsx", engine='openpyxl', sheet_name=n)
    # for i, row in df.iterrows():
    #     print("{}\t{}\n".format(i, row))
    
    min_date = df['Start'].min()
    
    max_date = df['End'].max() + timedelta(days=10)
    
    size = max_date - min_date
    size = int(size.days)
    
    arr = [[0 for i in range(size*2)] for i in range(df.shape[0]*2)]
    
    # for irow in range(len(arr)):
    #     for icol in range(size + 1):
    #         arr[irow][icol] = " "
    
    timeline_position = int(len(arr) / 2)
    
    
    start_dt = min_date
    end_dt = max_date
    col = 0
    for dt in daterange(start_dt, end_dt):
        arr[timeline_position][col] = dt
        col += 1
        # print(dt.strftime("%Y-%m-%d"))
    
    identify = {}
    for i in range(df.shape[0]):
        try:
            event = df.iloc[i]
            if event[1] not in identify:
                identify[event[1]] = event[2]
            start_position = int((event[3] - min_date).days)
            end_position = int((event[4] - min_date).days)
            draw_event(arr, event[1], start_position, end_position, timeline_position)
        except:
            pass
    
        
    
    worksheet = workbook.add_worksheet(name=n)
    
    merge_format = workbook.add_format({
        'bold':     True,
        'border':   6,
        'align':    'center',
        'valign':   'vcenter',
        'fg_color': '#D7E4BC',
        'size': 20
    })
    
    blank_format = workbook.add_format({
        'fg_color': '#263275',
    })
    
    for irow in range(len(arr)):
        if irow == timeline_position:
            continue
        end_col = 0
        value = ''
        icol = 0
        # for icol in range(size):
        while icol < size:
            # worksheet.write(irow, icol, " " if arr[irow][icol] == 0 else arr[irow][icol])
            if arr[irow][icol] == 0:
                worksheet.write(irow, icol, " ", blank_format)
                icol += 1
            else:
                value = arr[irow][icol]
                count = 0
                for j in range(icol+1, size):
                    if arr[irow][j] == value:
                        count += 1
                    else:
                        break
    
                if count > 0:
                    worksheet.merge_range(irow, icol, irow, icol + count, value, get_format(identify[value]))
                    icol += count + 1
                else:
                    worksheet.write(irow, icol, arr[irow][icol], get_format(identify[value]))
                    icol += 1
    
    date_format = workbook.add_format({
        'num_format': 'd/M',
        'bold':     True,
        'border':   6,
        'align':    'center',
        'valign':   'vcenter',
        'fg_color': '#ffff99',
        'size': 20
    })
    
    for icol in range(size):
        worksheet.write_datetime(timeline_position, icol, arr[timeline_position][icol], date_format)
    worksheet.set_zoom(50)
    
workbook.close()

