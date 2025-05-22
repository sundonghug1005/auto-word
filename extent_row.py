data = [["홍길동", "30", "서울"], ["김철수", "28", "부산"]]

table = doc.tables[0]

for row_data in data:
    row = table.add_row()
    for i, value in enumerate(row_data):
        row.cells[i].text = value
