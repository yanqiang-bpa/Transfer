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


import re
print(re.match( r'\d{8}', '20150623')).group(0)