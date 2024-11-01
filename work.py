from xlsxwriter import Workbook
from tkinter import Tk
from tkinter.filedialog import askopenfilenames
import pandas as pd

root = Tk()
root.withdraw()  # Tkinter 창을 숨김
 

night = ["18:00","18:30","20:00","20:30"]

# 여러 개의 파일을 선택할 수 있도록 설정
file_paths = askopenfilenames(title="엑셀 파일 선택", filetypes=[("Excel 파일", "*.xlsx")])

# 각 파일에 대해 작업 수행
for file_path in file_paths:
    df_data = pd.read_excel(file_path).fillna(' ')  # 각 파일별로 데이터프레임 읽어오기

    # 예시로 사용할 시설명을 첫 번째 행 == 0  '시설명' 열에서 추출
    Facility = df_data.loc[0, '시설명'] if '시설명' in df_data.columns else " "
    Enter_your_home_addres = df_data.loc[0, '주소'] if '주소' in df_data.columns else " "
    Name = df_data.loc[0, '예약자'] if '예약자' in df_data.columns else " "
    Phone_Num= df_data.loc[0, '휴대폰번호'] if '휴대폰번호' in df_data.columns else " "
    BirthDay = df_data.loc[0, '사업자등록번호'] if '사업자등록번호' in df_data.columns else ""
    Answer = df_data.loc[0, '대관목적'] if '대관목적' in df_data.columns else " "
    Name_of_event = df_data.loc[0, '행사명'] if '행사명' in df_data.columns else " "
    How_many_People = df_data.loc[0, '예상 인원'] if '예상 인원' in df_data.columns else " "
    Date = df_data.loc[0, '대관일'] if '대관일' in df_data.columns else " "
    start_time = df_data.loc[0, '시작시간'] if '시작시간' in df_data.columns else " "
    end_time = df_data.loc[0, '종료시간'] if '종료시간' in df_data.columns else " "
    Money = df_data.loc[0, '총결제금액'] if '총결제금액' in df_data.columns else " "
    
    Date = Date.replace("-", "년 ", 1).replace("-", "월 ", 1) + "일"

    # 엑셀 파일 생성
    wb = Workbook('homework.xlsx')
    ws = wb.add_worksheet()

    # 테두리 스타일 설정
    border_format = wb.add_format({
        'border': 1,
        'border_color': 'black',
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
    })

    border_format_left= wb.add_format({
        'border': 1,
        'border_color': 'black',
        'align': 'left',
        'valign': 'vcenter',
        'text_wrap': True,
    })

    ws.set_margins(top=1)

    # 셀 설정
    ws.set_column("A:A", 14)
    ws.set_column('B:B', 7)
    ws.set_column('C:C', 8)
    ws.set_column('D:D', 14)
    ws.set_column('E:E', 13)
    ws.set_column('F:F', 5)
    ws.set_column('G:G', 5)
    ws.set_column('H:H', 5)
    ws.set_column('I:I', 5)

    # 관람권 쪽은 작게 바꾸고 제목은 더 크게
    for i in range(0, 26):
        if i == 2:
            ws.set_row(i, 37) 
        else:
            ws.set_row(i, 27.7)

    # header
    ws.set_header('&B&C&20 체육시설 사용허가 신청서')

    # 데이터 입력 (신청인 정보 예시)
    ws.merge_range('A1:A3', "신청인", border_format)
    ws.merge_range('B1:C1', "주소", border_format)
    ws.merge_range('D1:I1', f"{Enter_your_home_addres}", border_format)
    ws.merge_range('B2:C3', "성명\n(법인이나 단체는\n법인명이나 단체대표자 명)", border_format)
    ws.merge_range('D2:D3', f"{Name}", border_format)
    ws.merge_range('E2:F2', "전화번호", border_format)
    ws.merge_range('G2:I2', f"{Phone_Num}", border_format)
    ws.merge_range('E3:F3', "생년월일\n(사업자등록번호)", border_format)
    ws.merge_range('G3:I3', f"{BirthDay}", border_format)

    ws.write('A4', "사용목적", border_format)
    ws.merge_range('B4:I4', f"{Answer}", border_format)

    # 시설명 및 사용내역
    ws.merge_range('A5:A6', "사용시설 명 \n및 사용내역", border_format)
    ws.merge_range('B5:D6', f"{Facility}", border_format)
    ws.merge_range('E5:F5', "중계방송", border_format)
    ws.merge_range('G5:I5', "  ", border_format)
    ws.merge_range('E6:F6', "기타", border_format)
    ws.merge_range('G6:I6', " ", border_format)

    # 기타 항목 추가
    ws.write('A7', "경기행사명", border_format)
    ws.merge_range('B7:D7', f"{Name_of_event}", border_format)
    ws.merge_range('E7:F7', "사용인원", border_format)
    ws.merge_range('G7:I7', f"{How_many_People}", border_format)

    # 계속하여 셀 병합 및 데이터 입력
    ws.merge_range('A8:A9', "사용기간", border_format)
    ws.write('B8', "주간", border_format)
    ws.write('B9', "야간", border_format)
    if start_time in night:
        ws.merge_range('C9:I9', f"{Date}  {start_time}부터   {Date}  {end_time}까지", border_format)
        ws.merge_range('C8:I8', " ",border_format)
        
    else:
        ws.merge_range('C8:I8', f"{Date}  {start_time}부터   {Date}  {end_time}까지", border_format)
        ws.merge_range('C9:I9', " ",border_format)
    
    ws.merge_range('A10:A11', "임원 또는\n지도요원", border_format)
    ws.merge_range('B10:I10', "임원                 ❨인❩            지도요원            ❨인❩          계              ❨인❩ ", border_format)
    ws.merge_range('B11:I11', "임원                 ❨인❩            지도요원            ❨인❩          계              ❨인❩ ", border_format)

    ws.merge_range('A12:A17', "관람권", border_format)
    ws.merge_range('B12:C12', " 관람권 종류 명 ", border_format)
    ws.write('D12', " 관람권 종류 명 ", border_format)
    ws.merge_range('E12:E17', "부속시설\n사용", border_format)
    ws.write('F12', "품목", border_format)
    ws.write('G12', "단가", border_format)
    ws.write('H12', "수량", border_format)
    ws.write('I12', "금액", border_format)

    for row in range(13, 18):
        ws.merge_range(f'B{row}:C{row}', "                 권", border_format)
        ws.write(f'D{row}', "     원       매", border_format)
        ws.write(f'F{row}', " ", border_format)
        ws.write(f'G{row}', " ", border_format)
        ws.write(f'H{row}', " ", border_format)
        ws.write(f'I{row}', " ", border_format)

    ws.write('A18', "기타", border_format)
    ws.merge_range('B18:I18', "", border_format)
    ws.write('A19', "사용료", border_format)
    ws.merge_range('B19:I19', f"금                           {Money}원", border_format)
    ws.merge_range('A20:I26', f"⌜안성시 체육시설 관리 운영 조례⌟ 제 3조에 따라 위와 같이 신청합니다.\n\n\n                                                  {Date}\n\n                                           신청인         {Name}        ❨인❩\n\n안성시설관리공단귀하\n\n 첨부: 1.사용❨사용❩계획서 1부                                     \n     2.폐기물 처리계획서 1부                                             ", border_format_left)

    # 파일 저장
    wb.close()
                                                            