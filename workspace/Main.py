from import_list import *

gl._init()
test_img_path = []
pic_name_save = []
pic_path = "C:/Users/m1861/Desktop/SCD/test_picture/"
pic_names = os.listdir(pic_path)  #图片路径
pic_num = 0  #用于统计文件数量
for pic_name in pic_names:
    pic_num = pic_num + 1
    if pic_name[-4:] == ".jpg" or pic_name[-4:] == ".png":
        test_img_path.append(pic_path + pic_name)
        pic_name_save.append(pic_name)

gl.set_value('pic_name_save', pic_name_save)
# print("finish_init!", flush=True)
OCR(test_img_path)
#results = gl.get_value('results')
#print(results)
#results_demo()
# time_start_write = time.time()
gl.set_value('picword', True)
gl.set_value('method', 'auto_first')
gl.set_value('model', 'speed_first')
gl.set_value('strategy', 'balance')
gl.set_value('frozen', True)
write_to_excel_all()
# print("time_write:", time.time() - time_start_write)

'''
for result in results:
        data = result['data']
        #save_path = result['save_path']
        for infomation in data:
            print('text: ', infomation['text'], '\nconfidence: ', infomation['confidence'], '\ntext_box_position: ', infomation['text_box_position'])
'''
