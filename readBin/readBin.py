import os
import glob
import xlrd
import xlwt

num = 115

def arithmetic(value):
        #print(value)
        result = {
                b'\x00': 0,
                b'\x01': 1,
                b'\x02': 2,
                b'\x03': 3,
                b'\x04': 4,
                b'\x05': 5,
                b'\x06': 6,
                b'\x07': 7,
                b'\x08': 8,
                b'\x09': 9,
                b'\x0A': 10,
                b'\x0B': 11,
                b'\x0C': 12,
                b'\x0D': 13,
                b'\x0E': 14,
                b'\x0F': 15,
                b'\x10': 16,
                b'\x11': 17,
                b'\x12': 18,
                b'\x13': 19,
                b'\x14': 20,
                b'\x15': 21,
                b'\x16': 22,
                b'\x17': 23,
                b'\x18': 24,
                b'\x19': 25,
                b'\x1A': 26,
                b'\x1B': 27,
                b'\x1C': 28,
                b'\x1D': 29,
                b'\x1E': 30,
                b'\x1F': 31,
                b'\x20': 32,
                b'\x21': 33,
                b'\x22': 34,
                b'\x23': 35,
                b'\x24': 36,
                b'\x25': 37,
                b'\x26': 38,
                b'\x27': 39,
                b'\x28': 40,
                b'\x29': 41,
                b'\x2A': 42,
                b'\x2B': 43,
                b'\x2C': 44,
                b'\x2D': 45,
                b'\x2E': 46,
                b'\x2F': 47,
                b'\x30': 48,
                b'\x31': 49,
                b'\x32': 50,
                b'\x33': 51,
                b'\x34': 52,
                b'\x35': 53,
                b'\x36': 53,
                b'\x37': 55,                
        }
        return result.get(value)

def read_bin(path):
	if(os.path.exists(path) != True):
		print("file no exist")
		return
	read_file = open(path,"rb+")
	wb = xlwt.Workbook()
	ws = wb.add_sheet("panduit")
	ws.write(0,0,'num')
	ws.write(0,1,'sku')
	ws.write(0,2,'pdutype')
	ws.write(0,3,'phasenum')
	ws.write(0,4,'phasetype')
	ws.write(0,5,'cbnum')
	ws.write(0,6,'outletnum')
	ws.write(0,7,'relaynum')
	global num
	for count in range(0,num):
		read_file.seek(count * 0x8B5,0)
		#if(count * 0x8B5) >= os.path.getsize(path):
		print("count:%d "%count,end='')
		#get sku
		data = read_file.read(16)
		
		str_value = ''
		for i in range(0,16):
			str_value += chr(data[i])
		ws.write(count+1,0,str(count+1)) #
		ws.write(count+1,1,str_value.strip()) #sku
		#pdutype
		read_file.seek(16,1)
		data = read_file.read(1)
		ws.write(count+1,2,str(arithmetic(data)))
		
		#phasenum
		data = read_file.read(1)
		ws.write(count+1,3,str(arithmetic(data)))	
		
		#phasetype
		data = read_file.read(1)
		print("phasetype:",end='')
		print(data,end='')
		print(" ",end='')
		ws.write(count+1,4,str(arithmetic(data)))
		
		#cbnum
		data = read_file.read(1)
		ws.write(count+1,5,str(arithmetic(data)))	
		
		#outletnum
		data = read_file.read(1)
		print("outletnum:",end='')
		print(data,end='')
		print(" ")
		ws.write(count+1,6,str(arithmetic(data)))
		
		#relaynum
		data = read_file.read(1)
		ws.write(count+1,7,str(arithmetic(data)))
	read_file.close()
	wb.save("test1.xls")

	
if __name__ == '__main__':
	path = 'skupanduit.bin'
	read_bin(path)
