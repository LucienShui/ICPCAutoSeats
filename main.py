import openpyxl
import random
width = 24
length = 16
tot = length * width
path = "xlsx.xlsx"


def get_coordinate(index):
	index = index - 1
	return index // width + 1, index % width + 1


def get_alpha(index):
	index = index - 1
	block_size = (width // 2) * (length // 2)
	belong = index // block_size
	return "%c%02d" % ("ABCD"[belong], index % block_size + 1)


def test():
	for i in range(1, tot + 1):
		x, y = get_coordinate(i)
		print("index = %d, (%d, %d), %s" % (i, x, y, get_alpha(i)))


def rand():
	return random.randint(1, tot)


def process():
	vis = []
	workbook = openpyxl.load_workbook(path)
	sheet = workbook[0]
	for row in sheet.rows:
		tmp = rand()
		while tmp in vis:
			tmp = rand()
		vis.append(tmp)
		row[0] = tmp


def min_dist(array):
	array_len = array.__len__()
	minimal = tot << 2
	if array_len == 1:
		return minimal
	for i in range(0, array_len):
		for j in range(i + 1, array_len):
			ux, uy = get_coordinate(array[i])
			vx, vy = get_coordinate(array[j])
			minimal = min(abs(ux - vx) + abs(uy - vy))
	return minimal


def check():
	workbook = openpyxl.load_workbook(path)
	sheet = workbook[0]
	sets = []
	cur = 0
	for row in sheet.rows:
		index = row[0]
		school_name = row[1]
		if school_name != cur:
			if min_dist(sets) <= 3:
				return False
		else:
			sets.append(index)
	if min_dist(sets) <= 3:
		return False
	return True


def show():
	workbook = openpyxl.load_workbook(path)
	sheet = workbook[0]
	for row in sheet.rows:
		for cell in row:
			print(cell, end="")
		print()
	

if __name__ == '__main__':
	show()
	print("Hello World!")
