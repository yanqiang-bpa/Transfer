# Test ssh key
# How to use the ISSUE in github
# dic = {
# 	"status" : 2,
# 	"samples" : [ 
# 			       {
# 			            "library_plate_id" : 1,
# 			            "library_plate": "abc"
# 			       }
# 			    ]
# }

# sample = {
# 			            "library_plate_id" : 2,
# 			            "library_plate": "dkfjd"
# 			       }

# dic["samples"].append(sample)

# print dic

# 正则表达式
import re
a = '20150623'
if(re.match( r'\d{8}$', '20150623')):
	a = a[:4] + '/' + a[4:6] + "/" + a[6:]
	print a
if(re.match(r'\d{4}.\d{2}.\d{2}$', '2015.12.23')):
	a = a.replace(".", "/")
	print(a)
if(re.match(r'\d{4}/\d{2}/\d{2}$', '2015/12/23')):
	a = a.replace(".", "/")
	print(a)

print	re.match(r'\d{4}/\d{2}/\d{2}$', '2015/12/23')
