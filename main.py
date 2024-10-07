import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Border, Side


# Function to get building data from user input
def get_building_data():
    buildings = []

    while True:
        building_name = int(input("동을 입력하세요 (종료하려면 0 입력): "))
        if building_name == 0:
            break
        while True:
            building_type = input("타입을 입력하세요 (해당 동 종료 시, end 입력): ")
            if building_type == "end":
                break
            room_num = int(input("라인을 입력하세요: "))
            highest_floor = int(input("최상층을 입력하세요: "))
            lowest_floor = int(input("최하층을 입력하세요: "))

            # Create list of floors
            floors = list(range(lowest_floor, highest_floor + 1))

            # Append the building information
            building = {
                "동": building_name,
                "타입": f"{building_type}",
                "층": floors,
                "호": room_num,
                "최저층": lowest_floor
            }
            buildings.append(building)

    return buildings


# Function to create nested data structure from building input
def create_nested_data(buildings):
    nested_data = {}

    for building in buildings:
        building_name = building["동"]
        building_type = building["타입"]
        room_num = building["호"]

        if building_name not in nested_data:
            nested_data[building_name] = {}

        if room_num not in nested_data[building_name]:
            nested_data[building_name][room_num] = {}

        if building_type not in nested_data[building_name][room_num]:
            nested_data[building_name][room_num][building_type] = []

        for floor in building["층"]:
            room_number = floor * 100 + room_num  # Generate room number by combining floor and room number
            nested_data[building_name][room_num][building_type].append(room_number)

    return nested_data

# Function to create the grid layout and save final Excel file
def create_grid_layout(input_filename, output_filename):
    # Load the workbook and sheet using openpyxl
    wb = openpyxl.load_workbook(input_filename, data_only=True)
    ws = wb.active
    set = wb.create_sheet(title="Grid", index=0)

    title = input("현장명을 입력하세요: ")
    set.cell(row=2, column=4).value = title

    # Create a nested dictionary to hold the room data
    nested_data = {}

    # 각 동의 최하층 및 최상층 계산
    buildings_info = {}
    min_floor = 1
    max_floor = 0

    for row in ws.iter_rows(min_row=3, values_only=True):  # Assuming the first row is a header
        building = row[0]  # Assuming '동' is the first column
        floor = row[3] // 100  # Assuming '호수' is the fifth column (e.g., 101 -> 1층)

        if building not in buildings_info:
            buildings_info[building] = {'lowest_floor': floor, 'highest_floor': floor}
        else:
            buildings_info[building]['lowest_floor'] = min(buildings_info[building]['lowest_floor'], floor)
            buildings_info[building]['highest_floor'] = max(buildings_info[building]['highest_floor'], floor)

    # 최하층과 최상층 계산
    for building, info in buildings_info.items():
        min_floor = min(min_floor, info['lowest_floor'])
        max_floor = max(max_floor, info['highest_floor'])

    nested_data = {}

    for row in ws.iter_rows(min_row=2, values_only=True):  # Assuming '동' is in the first column, '호' in third
        building_name = row[0]  # '동'
        building_type = row[1]  # '라인'
        line = row[2]  # '타입'
        room_number = row[3]  # '호수'
        floor_attribute = row[4]

        if building_name not in nested_data:
            nested_data[building_name] = {}

        # 호수가 존재하지 않는 경우 새로 추가
        if line not in nested_data[building_name]:
            nested_data[building_name][line] = {}

        # 타입이 존재하지 않는 경우 새로 추가
        if building_type not in nested_data[building_name][line]:
            nested_data[building_name][line][building_type] = []

        nested_data[building_name][line][building_type].append(room_number)

    # 현재 층 정보를 내림차순으로 정렬하여 1열에 작성
    current_row = 4
    val_list = range(max_floor, min_floor - 1, -1)  # Floors in descending order

    for i, val in enumerate(val_list):
        set.cell(row=current_row, column=1, value=val)
        current_row += 1

    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    current_col = 3
    max_floor += 3
    # 동, 라인 번호 우선으로 정렬하여 출력
    for building_name in sorted(nested_data.keys()):  # 동 이름으로 정렬
        lines = nested_data[building_name]
        save_col = current_col

        for line in sorted(lines.keys()):  # 라인 번호로 정렬
            room_types = lines[line]
            set.cell(row=max_floor + 3, column=current_col + 1, value=building_name).alignment = Alignment(
                horizontal='center')
            set.cell(row=max_floor + 3, column=current_col + 1, value=building_name).border = thin_border


            # 타입 우선 정렬을 제거했습니다.
            for building_type in room_types.keys():
                # 층 정보를 내림차순으로 나열
                set.cell(row=max_floor + 2, column=current_col + 1, value=building_type).alignment = Alignment(
                    horizontal='center')
                set.cell(row=max_floor + 2, column=current_col + 1, value=building_type).border = thin_border

                # 방 번호를 층 정보를 기준으로 출력
                for room_num in room_types[building_type]:
                    floor = room_num // 100  # 층 정보
                    room_row = max_floor - floor + 1  # 층에 따른 행 번호 조정
                    set.cell(row=room_row, column=current_col + 1, value=room_num).alignment = Alignment(
                        horizontal='right')
                    set.cell(row=room_row, column=current_col + 1, value=room_num).border = thin_border

                current_col += 1

        set.merge_cells(start_row=max_floor + 3, start_column=save_col + 1, end_row=max_floor + 3, end_column=current_col)
        set.cell(row=max_floor + 3, column=save_col + 1).border = thin_border

        current_col += 2


    wb.save(output_filename)
    print(f"Final layout saved to {output_filename}")


def save_to_excel(nested_data, filename):
    wb = openpyxl.Workbook()
    ws = wb.active

    # Column headers
    ws.cell(row=1, column=1, value="동")
    ws.cell(row=1, column=2, value="라인")
    ws.cell(row=1, column=3, value="타입")
    ws.cell(row=1, column=4, value="호수")
    ws.cell(row=1, column=5, value="층 속성")

    current_row = 2

    for building_name in sorted(nested_data.keys()):  # 동 이름으로 정렬
        for line in sorted(nested_data[building_name].keys()):  # 라인 번호로 정렬
            for building_type in nested_data[building_name][line].keys():  # 타입으로 정렬


                room_numbers = nested_data[building_name][line][building_type]
                lowest_floor = min([room_num // 100 for room_num in room_numbers])

                for room_num in nested_data[building_name][line][building_type]:
                    # 층 속성 결정
                    if room_num // 100 == 1:
                        floor_attribute = "최하층"
                    elif room_num // 100 == lowest_floor:
                        floor_attribute = "피트층"
                    else:
                        floor_attribute = "기준층"

                    ws.cell(row=current_row, column=1, value=building_name)
                    ws.cell(row=current_row, column=2, value=line)
                    ws.cell(row=current_row, column=3, value=building_type)
                    ws.cell(row=current_row, column=4, value=room_num)
                    ws.cell(row=current_row, column=5, value=floor_attribute)

                    current_row += 1

    wb.save(filename)
    print(f"Data saved to {filename}")


# Main function to integrate all parts
def main():
    # Step 1: Get building data from user
    buildings = get_building_data()

    # Step 2: Create nested data
    nested_data = create_nested_data(buildings)

    # Step 3: Save to intermediate Excel file
    intermediate_file = "building_room_dataaaa.xlsx"
    save_to_excel(nested_data, intermediate_file)
    # Step 4: Create grid layout and save final Excel file
    output_file = input("저장하고자하는 최종 파일명을 입력하세요 : ")
    output_file = "save_to_excel/"+ output_file + ".xlsx"
    create_grid_layout(intermediate_file, output_file)


if __name__ == "__main__":
    main()
