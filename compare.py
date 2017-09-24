from hash import File_hash
import sys
def file_md5_cp():
	#print(sys.argv[1])
	new=File_hash(sys.argv[1])
	#print(new.file_p)
	#new.__init__()
	md5=new.get_hash()
	print(md5)

	new2=File_hash(sys.argv[2])
	#print(new.file_p)
	#new.__init__()
	md5_2=new2.get_hash()
	print(md5_2)

	if md5 ==  md5_2:
		print('same file')
	else:
		print('diffrent file')


if __name__ == "__main__":
	file_md5_cp()
	pass