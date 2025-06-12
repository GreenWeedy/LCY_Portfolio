import argparse
import serial
from openpyxl import Workbook


def read_serial_to_excel(port, baudrate, output, count):
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
    parser.add_argument('port', help='Serial port (e.g., COM3 or /dev/ttyUSB0)')
    parser.add_argument('-b', '--baudrate', type=int, default=9600, help='Baud rate')
    parser.add_argument('-o', '--output', default='output.xlsx', help='Output Excel file')
    parser.add_argument('-c', '--count', type=int, default=100, help='Number of lines to read')
    args = parser.parse_args()
    read_serial_to_excel(args.port, args.baudrate, args.output, args.count)


if __name__ == '__main__':
    main()
