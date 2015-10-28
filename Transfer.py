# coding=utf-8
import xlrd
import datetime
import os
import sys  
import json
import re
from imp import reload
reload(sys)  
# sys.setdefaultencoding('utf-8')

# command = "mongoexport -d sample_track_dev -c tasks -o C:\Users\yanqiang\Desktop\\test.dat"
# os.system(command)

# [u'\u4efb\u52a1\u5355\u540d\u79f0', 					0 任务单名称
# u'\u5f52\u5c5e', 										1 归属
# u'\u603b\u9879\u76ee\u540d\u79f0', 					2 总项目名称
# u'\u603b\u9879\u76ee\u4ee3\u7801', 					3 总项目代码
# u'\u5b50\u9879\u76ee\u540d\u79f0', 					4 子项目名称
# u'\u5b50\u9879\u76ee\u4ee3\u7801', 					5 子项目代码
# u'\u76f8\u5173\u8d1f\u8d23\u4eba\u53ca\u90ae\u7bb1', 	6 相关负责人及邮箱
# u'\u4fe1\u606f\u8d1f\u8d23\u4ebaCGI\u8d26\u53f7', 	7 信息负责人CGI帐号
# u'\u5f00\u59cb\u65e5\u671f', 							8 开始日期
# u'\u622a\u6b62\u65e5\u671f', 							9 截至日期
# u'\u6837\u54c1\u540d\u79f0*', 						10 样品名称*
# u'\u6837\u54c1\u7f16\u53f7/\u6587\u5e93\u7f16\u53f7*', 11 样品编号/文库编号*
# u'\u6837\u54c1\u7c7b\u578b*', 						12 样品类型*
# u'\u6587\u5e93\u7c7b\u578b*', 						13 文库类型*
# u'\u5efa\u5e93Adaptor', 								14 建库Adaptor
# u'\u82af\u7247\u540d\u79f0', 							15 芯片名称
# u'\u6742\u4ea4\u57fa\u6570', 							16杂交基数
# u'\u6742\u4ea4\u524dPool/\u6742\u4ea4\u540ePool', 	17杂交前Pool/杂交后Pool
# u'\u7269\u79cd*', 									18物种*
# u'\u5efa\u5e93\u4e2a\u6570*', 						19建库个数*
# u'\u6d4b\u5e8f\u7c7b\u578b*', 						20测序类型*
# u'\u6d4b\u5e8fAnchor', 								21测序Anchor
# u'\u539f\u59cb\u6570\u636e\u91cf\uff08Gbp)*', 		22原始数据量（Gbp)*
# u'\u6837\u54c1\u4f53\u79ef(ul\uff09*', 				23样品体积(ul）*
# u'\u6837\u54c1\u6d53\u5ea6\uff08ng/ul)*', 			24样品浓度（ng/ul)*
# u'\u5907\u6ce8', 										25备注
# u'Sample', 											26Sample
# u'comment', 											27comment
# u'lane', 												28lane
# u'coverage', 											29coverage
# u'\u5efa\u5e93\u5f00\u59cb', 							30建库开始
# u'\u5efa\u5e93\u5b8c\u6210', 							31建库完成
# u'\u4e0a\u673a\u65e5\u671f', 							32上机日期
# u'\u4e0b\u673a\u65e5\u671f']							33下机日期
# u'\u4efb\u52a1\u5355\u7c7b\u578b'						34任务单类型
# u'\u6570\u636e\u5206\u6790\u7c7b\u578b'				35数据分析类型

# 2Ad 16，288，5017~5021，7092，这些行，没有任务单名称和总项目名称
# 2829，3729行，样品名称列填的是1/8(2014/1/8)

__s_date = datetime.date (1899, 12, 31).toordinal() - 1
def getdate(date):
	try:
	    if isinstance(date , float ):
	        date = int(date )
	    d = datetime.date .fromordinal(__s_date + date )
	    return d.strftime("%Y-%m-%d")
	except TypeError:
		return str(date)

data = xlrd.open_workbook(u'BB任务单.xlsx')

table = data.sheets()[0]

nrows = table.nrows

lastTaskName = ""
lastProjName = ""

taskDic = {
    "taskList_name" : "",
    "division" : "",
    "project_name" : "",
    "project_code" : "",
    "subproject_name" : "",
    "pm_name" : "",
    "pm_email" : "",
    "experiment_user" : "",
    "experiment_group" : "",
    "cgichina_account" : "",
    "start_date" : "",
    "end_date" : "",
    "task_type_id" : "",
    "task_type" : "",
    "task_library_type" : "",
    "species" : "",
    "specificSpecies" : "",
    "library_email_group" : "",
    "meta" : {
        "update_user" : "",
        "create_user" : "",
        "create_date" : "",
        "update_date" : ""
    },
    "status" : "",
    "samples" : [ 
        {
            "library_plate_id" : None,
            "library_plate" : "",
            "meta" : {
                "create_date" : "",
                "update_date" : ""
            },
            "comment" : "",
            "sample_concentration" : "",
            "sample_volume" : "",
            "analysis_type" : "",
            "lane_count" : "",
            "clean_data_size" : "",
            "sequencing_anchor" : "",
            "sequencing_type" : "",
            "library_count" : "",
            "specificSpecies" : "",
            "species" : "",
            "pooling_order" : "",
            "pooling_base" : "",
            "chip_name" : "",
            "library_adaptor" : "",
            "library_type" : "",
            "sample_code" : "",
            "sample_library_name" : "",
            "task" : "",
            "task_id" : None,
            "column_index" : "",
        }
    ]
}
sampleDic = {
    "library_plate_id" : None,
    "library_plate" : "",
    "meta" : {
        "create_date" : "",
        "update_date" : ""
    },
    "comment" : "",
    "sample_concentration" : "",
    "sample_volume" : "",
    "analysis_type" : "",
    "lane_count" : "",
    "clean_data_size" : "",
    "sequencing_anchor" : "",
    "sequencing_type" : "",
    "library_count" : "",
    "specificSpecies" : "",
    "species" : "",
    "pooling_order" : "",
    "pooling_base" : "",
    "chip_name" : "",
    "library_adaptor" : "",
    "library_type" : "",
    "sample_code" : "",
    "sample_library_name" : "",
    "task" : "",
    "task_id" : "",
    "column_index" : ""
}

for i in range(nrows):
	# 每一行数据的列表
	sample = table.row_values(i)
	# if()
	# print getdate(sample[-1])
	# print sample
	
	taskName = sample[0]
	projectName = sample[2]

	# 如果这一行是数据行（即不是表头，又不为空行）
	if((taskName !="" or projectName!="") and not projectName.startswith(u'\u603b\u9879\u76ee\u540d\u79f0')):
		# print sample
		print(i)

		# 即和上一个属于同一个task
		if(taskName == lastTaskName and projectName == lastProjName):
			# 生成一个新的sample数据
			sampleDic = {
	            "library_plate_id" : None,
	            "library_plate" : "",
	            "meta" : {
	                "create_date" : "",
	                "update_date" : ""
	            },
	            "__v" : 0,
	            "comment" : sample[27],
	            "sample_concentration" : sample[24],
	            "sample_volume" : sample[23],
	            "analysis_type" : sample[35],
	            "lane_count" : sample[28],
	            "clean_data_size" : sample[22],
	            "sequencing_anchor" : "141",
	            "sequencing_type" : "",
	            "library_count" : 1,
	            "specificSpecies" : "",
	            "species" : "2",
	            "pooling_order" : "false",
	            "pooling_base" : sample[16],
	            "chip_name" : "",
	            "library_adaptor" : "",
	            "library_type" : "6",
	            "sample_code" : sample[11],
	            "sample_library_name" : sample[10],
	            "task" : taskName,
	            "task_id" : "",
	            "column_index" : 1
	        }
			taskDic["samples"].append(sampleDic)
		# 一个新的task
		else:
			if(taskDic["taskList_name"] !="" or taskDic["project_name"] != ""):
				# print taskDic
				taskJson = json.dumps(taskDic)
				print(taskJson)
				#Windows
				f = open("C:\Users\yanqiang\Desktop\out.dat","w")
				f.write(taskJson)
				f.close()
				command = "mongoimport -d sample_track_dev -c tasks C:\Users\yanqiang\Desktop\out.dat"
				os.system(command)

				# Linux
				# f = open("/home/tony/out.dat","w")
				# f.write(taskJson)
				# f.close()
				# command = "mongoimport -d sample_track_dev -c tasks /home/tony/out.dat"
				# os.system(command)
			

			if(sample[8] == ""):
				start_date = ""
			else:
				
				if(re.match( r'\d{8}.0$', str(sample[8]))):
					sample[8] = str(sample[8])[:-2]
					sample[8] = sample[8][:4] + '-' + sample[8][4:6] + "-" + sample[8][6:]
					start_date = sample[8]
				if(re.match(r'\d{4}.\d{2}.\d{2}$', str(sample[8]))):
					start_date = sample[8].replace(".", "-")
				else:
					start_date = getdate(sample[8])

			if(sample[9] == ""):
				end_date = ""
			else:
				
				if(re.match( r'\d{8}.0$', str(sample[9]))):
					sample[9] = str(sample[9])[:-2]
					sample[9] = sample[9][:4] + '-' + sample[9][4:6] + "-" + sample[9][6:]
					end_date = sample[9]
				if(re.match(r'\d{4}.\d{2}.\d{2}$', str(sample[9]))):
					end_date = sample[9].replace(".", "-")
				else:
					end_date = getdate(sample[9])

			taskDic = {
			    "taskList_name" : taskName,
			    "division" : sample[1],
			    "project_name" : sample[2],
			    "project_code" : sample[3],
			    "subproject_name" : sample[4],
			    "pm_name" : sample[6],
			    "pm_email" : "",
			    "experiment_user" : "",
			    "experiment_group" : "",
			    "cgichina_account" : sample[7],
			    "start_date" : start_date,
			    "end_date" : end_date,
			    "task_type_id" : "",
			    "task_type" : sample[34],
			    "task_library_type" : "",
			    "species" : sample[18],
			    "specificSpecies" : "",
			    "library_email_group" : "",
			    "comment" : sample[27],
			    "meta" : {
			        "update_user" : "",
			        "create_user" : "",
			        "create_date" : "",
			        "update_date" : ""
			    },
			    "status" : 2,
			    "samples" : [ 
			        {
			            "library_plate_id" : "",
			            "library_plate" : "",
			            "meta" : {
			                "create_date" : "",
			                "update_date" : ""
			            },
			            "__v" : 0,
			            "comment" : sample[25],
			            "sample_concentration" : sample[24],
			            "sample_volume" : sample[23],
			            "analysis_type" : sample[35],
			            "lane_count" : sample[28],
			            "clean_data_size" : sample[22],
			            "sequencing_anchor" : sample[21],
			            "sequencing_type" : sample[20],
			            "library_count" : sample[19],
			            "specificSpecies" : "",
			            "species" : sample[18],
			            "pooling_order" : sample[17],
			            "pooling_base" : sample[16],
			            "chip_name" : sample[15],
			            "library_adaptor" : sample[14],
			            "library_type" : sample[13],
			            "sample_code" : sample[11],
			            "sample_library_name" : sample[10],
			            "task" : taskName,
			            "task_id" : "",
			            "column_index" : 1
			        }
			    ],
			    "__v" : 0
			}

		lastTaskName = taskName;
		lastProjName = projectName;
	elif(i>7510 and taskDic["taskList_name"] !="" and taskDic["project_name"] != ""):
		# print taskDic
		taskJson = json.dumps(taskDic)
		print(taskJson)
		#Windows
		f = open("C:\Users\yanqiang\Desktop\out.dat","w")
		f.write(taskJson)
		f.close()
		command = "mongoimport -d sample_track_dev -c tasks C:\Users\yanqiang\Desktop\out.dat"
		os.system(command)

		# Linux
		# f = open("/home/tony/out.dat","w")
		# f.write(taskJson)
		# f.close()
		# command = "mongoimport -d sample_track_dev -c tasks /home/tony/out.dat"
		# os.system(command)

		break;
