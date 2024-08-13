import os
import re
import pandas as pd
from docx import Document


class To_excel:
	def __init__(self):
		self.tables_data_first_four = []  # 用于存放前四行的数据
		self.tables_data_fifth_row = []  # 用于存放第五行的数据
		self.tables_data_sixth_row = []  # 用于存放第五行的数据
		self.big_dicts = {}
		self.phone = None
		self.native_place = None
		self.address = None
		self.path_files = None
		self.body = None
		self.dit_lis = []
	def read_world(self, file_path):
		"""
		主要是读取world文档的
		:param file_path: 文件路径
		:return:
		"""
		doc = Document(file_path)
		for table in doc.tables:
			table_data_first_four = []
			table_data_fifth_row = []
			table_data_sixfh_row = []
			for i, row in enumerate(table.rows):
				# 获取前四行数据
				if i < 4:
					row_data = [cell.text for cell in row.cells]
					table_data_first_four.append(row_data)
				elif i == 4:
					table_data_fifth_row = [cell.text for cell in row.cells]
				# 提取第六行数据
				elif i == 5:
					table_data_sixfh_row = [cell.text for cell in row.cells]
					break  # 如果已经到第五行，则可以退出循环，因为我们不再关心之后的行
			# 将提取的数据存储到对应的列表中
			self.tables_data_first_four.append(table_data_first_four)
			if table_data_fifth_row:  # 仅当第五行存在时
				self.tables_data_fifth_row.append(table_data_fifth_row)
			if table_data_sixfh_row:
				self.tables_data_sixth_row.append(table_data_sixfh_row)
	
	def save_dict(self, file_path):
		"""
		主要是拼装内容的
		:param file_path: world文档路径，
		:return: self.body 是每次拼装成的字典
		"""
		self.read_world(file_path)
		# 处理前四行
		list1 = []
		for table in self.tables_data_first_four:
			for i in table:
				seen = set()
				set_i = [x for x in i if not (x in seen or seen.add(x))]
				# print(set_i)
				list1.append(set_i)
		merged_list = [item for sublist in list1 if len(sublist) % 2 == 0 for item in sublist]
		result_dicts = [{merged_list[i]: merged_list[i + 1]} for i in range(0, len(merged_list), 2)]
		# print(result_dicts)
		
		for dits in result_dicts:
			self.big_dicts.update(dits)
		# print(self.big_dicts)
		# 处理第五行数据
		for table_data in self.tables_data_fifth_row:
			for i in set(table_data):
				# print(i)
				if i.isdigit():
					# pass
					self.phone = i
				if '地址' in i:
					self.native_place = i.split("：")[1]
		# 处理第六行
		for table_data in self.tables_data_sixth_row:
			for i in set(table_data):
				if "@" in i:
					email_data = i
				if '地址' in i:
					self.address = i.split("：")[1]
		# 处理下银行卡区分银行卡和开户行
		bank = self.big_dicts.get('中国银行\n卡号及开户行')
		bank_address = re.findall(r'[\u4e00-\u9fa5]', bank)  # 获取汉字
		str_bank_address = re.sub(r'\s', '', (' '.join(bank_address)))  # 获取的汉字是list转str.并去掉转后的空格
		bank_id = re.findall(r'\d+', bank)
		str_bank_id = ' '.join(bank_id)
		self.body = {
			'员工姓名': self.big_dicts.get("姓名"),
			'学校': self.big_dicts.get("毕业院校"),
			'专业': self.big_dicts.get("专业"),
			'婚姻状况': self.big_dicts.get("婚姻状况"),
			'住址': self.address,
			'党员身份': self.big_dicts.get("党团员"),
			'户籍所在地': self.big_dicts.get("籍贯\n（省/市/县）"),
			'身份证号': self.big_dicts.get('身份证号码'),
			'工资卡号': str_bank_id,
			'开户行': str_bank_address,
			'手机号码': self.phone
		}

		return self.body
	
	def get_all_path(self, s):
		# 当前目录的文件 s为格式内容
		"""
		:param s: 格式的内容 s可以填上'.docx'/'.xlsx'
		:return: 返回所有文件路径
		"""
		current_path = os.getcwd()
		files_and_dirs = os.listdir(current_path)
		self.path_files = [os.path.join(current_path, file) for file in files_and_dirs if file.endswith(s)]
		
		return self.path_files
	
	def all_world_to_dict(self):
		"""
		整理body的数据把这些数据变成需要存储成excel需要的类型
		:return:
		"""
		
		self.get_all_path('.docx')
		for i in self.path_files:
			body = self.save_dict(i)
			self.dit_lis.append(body)
		print(self.dit_lis)
		dic = {}
		for i in self.dit_lis:
			for key, value in i.items():
				if key in dic:
					dic[key].append(value)
				else:
					dic[key] = [value]
		print(dic)
		return dic
	
	def save_excel(self):
		"""
		把整理的数据存到Excel表里
		:return:
		"""
		xlsx_path = self.get_all_path('.xlsx')[0]
		data = self.all_world_to_dict()
		pf = pd.DataFrame(data)
		save_path = xlsx_path
		pf.to_excel(save_path, index=False)


if __name__ == '__main__':
	a = To_excel()
	a.save_excel()
