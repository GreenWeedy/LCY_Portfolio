import argparse
import serial
from openpyxl import Workbook


def read_serial_to_excel(port, baudrate, output, count):
    # timeout은 한 줄을 읽을 때 대기할 최대 시간(초)으로 필요에 따라 변경 가능
    ser = serial.Serial(port, baudrate, timeout=1)
    wb = Workbook()
    ws = wb.active
    ws.append(["Data"])  # header
    try:
        for _ in range(count):
            line = ser.readline().decode('utf-8', errors='ignore').strip()
            if line:
                ws.append([line])
    finally:
        ser.close()
        wb.save(output)


def main():
    parser = argparse.ArgumentParser(description="Read serial data and save to Excel")
    parser.add_argument('port', help='Serial port (e.g., COM3 or /dev/ttyUSB0)')  # 변경 시 데이터를 읽어올 포트가 달라집니다
    parser.add_argument('-b', '--baudrate', type=int, default=9600, help='Baud rate')  # 장치와의 통신 속도, 장치 설정과 맞지 않으면 통신이 실패할 수 있습니다
    parser.add_argument('-o', '--output', default='output.xlsx', help='Output Excel file')  # 결과를 저장할 파일 이름으로 변경 가능
    parser.add_argument('-c', '--count', type=int, default=100, help='Number of lines to read')  # 몇 줄을 읽을지 지정하며 값을 바꾸면 읽는 양이 달라집니다
    args = parser.parse_args()
    read_serial_to_excel(args.port, args.baudrate, args.output, args.count)


if __name__ == '__main__':
    main()
