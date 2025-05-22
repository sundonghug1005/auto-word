target_text = "##이름##"  # 템플릿에 미리 넣어둔 식별자

for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            if target_text in cell.text:
                cell.text = cell.text.replace(target_text, "홍길동")
