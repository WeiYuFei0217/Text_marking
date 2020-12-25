from import_list import *
import paddlehub as hub

def OCR(test_img_path):
    # 加载移动端预训练模型
    # time_start_load = time.time()
    print("是否保存为图片：", "True" if gl.get_value('picword')==True else "False", flush=True)
    print("置信度阈值", "50%" if gl.get_value('strategy')=='balance' else \
            ("10%" if gl.get_value('strategy')=='wide_first' else "90%"), flush=True)
    if gl.get_value('model')=='speed_first':
        print("加载mobile模型", flush=True)
        ocr = hub.Module(name="chinese_ocr_db_crnn_mobile")
    else:
        print("加载server模型", flush=True)
        ocr = hub.Module(name="chinese_ocr_db_crnn_server") #服务器端模型
    # 读取测试文件夹test.txt中的照片路径
    # print("time_load:", time.time() - time_start_load, flush=True)
    # time_start_readpic = time.time()
    np_images =[cv2.imread(image_path) for image_path in test_img_path]
    # print("time_read_picture", time.time() - time_start_readpic, flush=True)
    a = ocr.recognize_text(
        images=np_images,         # 图片数据，ndarray.shape 为 [H, W, C]，BGR格式;
        use_gpu=False,            # 是否使用 GPU；若使用GPU，请先设置CUDA_VISIBLE_DEVICES环境变量
        output_dir='../Output_Picture',      # 图片的保存路径，默认设为 ocr_result;
        visualization=True if gl.get_value('picword')==True else False,       # 是否将识别结果保存为图片文件;
        box_thresh=0.5 if gl.get_value('strategy')=='balance' else \
            (0.1 if gl.get_value('strategy')=='wide_first' else 0.9),           # 检测文本框置信度的阈值;
        text_thresh=0.5 if gl.get_value('strategy')=='balance' else \
            (0.1 if gl.get_value('strategy')=='wide_first' else 0.9))          # 识别中文文本置信度的阈值;
    '''
    print(type(a))
    print(a)
    file=open('C:/Users/m1861/Desktop/software curriculum design/data.txt','w') 
    file.write(str(a))
    file.close()
    '''
    gl.set_value('results', a)