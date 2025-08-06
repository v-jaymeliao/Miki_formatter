from docx import Document
from datetime import timedelta
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Function to convert time format [h]:mm to minutes
def time_to_minutes(time_str):
    if not time_str:
        return 0
    hours, minutes = map(int, time_str.split(':'))
    return hours * 60 + minutes

# Function to convert minutes back to [h]:mm format
def minutes_to_time(minutes):
    hours = minutes // 60
    minutes = minutes % 60
    return f"{hours}:{minutes:02d}"

# Function to sum time values in the 'Used' column
def sum_times(time_list):
    total_minutes = sum(time_to_minutes(time) for time in time_list)
    return minutes_to_time(total_minutes)

# Function to create or modify the table in Word
def edit_word_table(doc_path):
    doc = Document(doc_path)
    table = None

    # Search for the table with specific columns
    for tbl in doc.tables:
        headers = [cell.text.strip() for cell in tbl.rows[0].cells]
        
        if 'Package' in headers and 'Service' in headers and 'Type' in headers:
            table = tbl
            break

    if not table:
        print("Table not found.")
        return

    # Extract column indices for easier access
    headers = [cell.text.strip() for cell in table.rows[0].cells]
    package_idx = headers.index('Package')
    service_idx = headers.index('Service')
    type_idx = headers.index('Type')
    purchased_idx = headers.index('Purchased')
    used_idx = headers.index('Used')
    remaining_idx = headers.index('Remaining')
    trend_idx = headers.index('Trend')

    # Get the current rows data for sum calculations
    purchased_values = []
    used_values = []
    for row in table.rows[1:]:  # Skip header row
        purchased_values.append(row.cells[purchased_idx].text.strip())
        used_values.append(row.cells[used_idx].text.strip())

    # Sum the 'Purchased' and 'Used' values
    total_purchased = sum(float(value) for value in purchased_values if value)
    total_used = sum_times(used_values)

    # Add a new row for the 'Total'
    new_row = table.add_row()
    new_row.cells[service_idx].text = 'Total'
    new_row.cells[purchased_idx].text = str(total_purchased)
    new_row.cells[used_idx].text = total_used
    new_row.cells[remaining_idx].text = ''
    new_row.cells[trend_idx].text = ''

    # Save the modified document
    doc.save('modified_' + doc_path)




def highlight_table_yellow(tbl):
    # 先標記顏色
    for row in tbl.rows:
        for cell in row.cells:
            cell._element.get_or_add_tcPr().append(
                parse_xml(r'<w:shd {} w:fill="FFFF00"/>'.format(nsdecls('w')))
            )
    # 取得欄位索引
    headers = [cell.text.strip() for cell in tbl.rows[0].cells]
    col_count = len(headers)
    # 檢查最後一列 Service 欄是否已經是 Total:
    service_idx = None
    for idx, h in enumerate(headers):
        if h == 'Service':
            service_idx = idx
            break
    if service_idx is not None and tbl.rows[-1].cells[service_idx].text.strip() == 'Total:':
        return  # 已經有 Total: 就不再新增

    # 準備加總用
    purchased_sum = 0.0
    used_sum = 0
    remaining_sum = 0
    purchased_idx = used_idx = remaining_idx = trend_idx = service_idx = None
    for idx, h in enumerate(headers):
        if h == 'Purchased':
            purchased_idx = idx
        if h == 'Used':
            used_idx = idx
        if h == 'Remaining':
            remaining_idx = idx
        if h == 'Trend':
            trend_idx = idx
        if h == 'Service':
            service_idx = idx
    # 加總 purchased, used, remaining
    for row in tbl.rows[1:]:
        if purchased_idx is not None:
            val = row.cells[purchased_idx].text.strip()
            if val:
                try:
                    purchased_sum += float(val)
                except:
                    pass
        if used_idx is not None:
            val = row.cells[used_idx].text.strip()
            if val:
                try:
                    h, m = map(int, val.split(':'))
                    used_sum += h*60 + m
                except:
                    pass
        if remaining_idx is not None:
            val = row.cells[remaining_idx].text.strip()
            if val:
                try:
                    h, m = map(int, val.split(':'))
                    remaining_sum += h*60 + m
                except:
                    pass
    # 新增一列，只填指定欄位
    new_row = tbl.add_row()
    for idx in range(col_count):
        header = headers[idx]
        cell = new_row.cells[idx]
        if header == 'Service':
            cell.text = 'Total:'
        elif header == 'Purchased':
            cell.text = f"{purchased_sum}"
        elif header == 'Used':
            h = used_sum // 60
            m = used_sum % 60
            cell.text = f"{h}:{m:02d}"
        elif header == 'Remaining':
            h = remaining_sum // 60
            m = remaining_sum % 60
            cell.text = f"{h}:{m:02d}"
        elif header == 'Trend':
            percent = 0.0
            if purchased_sum > 0:
                percent = used_sum * 24 / purchased_sum 
            cell.text = f"{percent:.2f}%"
        else:
            cell.text = ''

def find_and_highlight_specific_table(doc_path):
    doc = Document(doc_path)
    target_headers = ["Package", "Service", "Type", "Purchased", "Used", "Remaining", "Trend"]

    def recursive_search(tbls):
        for tbl in tbls:
            if len(tbl.rows) > 0:
                headers = [cell.text.strip() for cell in tbl.rows[0].cells]
                if headers == target_headers:
                    highlight_table_yellow(tbl)
            # 遞迴搜尋巢狀表格
            for row in tbl.rows:
                for cell in row.cells:
                    if cell.tables:
                        recursive_search(cell.tables)
    recursive_search(doc.tables)
    outname = 'findtable' + doc_path
    doc.save(outname)
    print(f"已將符合條件的表格標記為黃色，並儲存為 {outname}")




find_and_highlight_specific_table('Support Services Customer Paginated Report.docx')