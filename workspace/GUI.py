from tkinter import *
from tkinter.messagebox import *
from tkinter.filedialog import *
from tkinter.font import nametofont
from tkinter import ttk
from import_list import *

global filename
global filedir
sys.setrecursionlimit(262144) # 设置栈的最大深度

class Application(Frame):

    def __init__(self, master=None):
        super().__init__(master) # super()代表父类定义而不是父类对象
        global Win_W, Win_H, Win_X, Win_Y
        Win_W = 1600 # 可自由调节
        Win_H = int(Win_W / 2)
        '''
        root.update_idletasks()
        Win_X = int((root.winfo_screenwidth() - root.winfo_reqwidth()) / 2)
        Win_Y = int((root.winfo_screenheight() - root.winfo_reqheight()) / 2)
        '''
        Win_X = 100
        Win_Y = int(Win_X / 2)
        master.geometry(str(Win_W)+"x"+str(Win_H)+"+"+str(Win_X)+"+"+str(Win_Y)) # 窗口WHXY
        self.master = master
        self.grid()
        self.createWidget()
        # self.ocr = hub.Module(name="chinese_ocr_db_crnn_mobile") 预加载OCR模型
        '''
        np_images =[cv2.imread("C:\\Users\\m1861\\Desktop\\SCD\\test_picture") for image_path in test_img_path] 
        a = ocr.recognize_text(
                            images=np_images,         # 图片数据，ndarray.shape 为 [H, W, C]，BGR格式;
                            use_gpu=False,            # 是否使用 GPU；若使用GPU，请先设置CUDA_VISIBLE_DEVICES环境变量
                            output_dir='../Output_Picture',      # 图片的保存路径，默认设为 ocr_result;
                            visualization=True,       # 是否将识别结果保存为图片文件;
                            box_thresh=0.5,           # 检测文本框置信度的阈值;
                            text_thresh=0.5)          # 识别中文文本置信度的阈值;
        '''
    
    def createWidget(self):
        """创建组件"""

        """菜单"""
        # 创建菜单栏
        menubar = Menu(self.master, tearoff=0)
        # 创建一级菜单（第一条竖线排列）
        menuFile = Menu(menubar, tearoff=0)
        menuLable = Menu(menubar, tearoff=0)
        menuOutput = Menu(menubar, tearoff=0)
        menuMethod = Menu(menubar, tearoff=0)
        menuView = Menu(menubar, tearoff=0)
        # 创建二级菜单（第二条竖线排列）（标注设置）
        menuLable_auto = Menu(menuLable, tearoff=0)
        menuLable_manual = Menu(menuLable, tearoff=0)
        # 创建二级菜单（第二条竖线排列）（输出设置）
        menuOutput_xml = Menu(menuOutput, tearoff=0)
        # 创建二级菜单（第二条竖线排列）（视图）
        menuView_check = Menu(menuView, tearoff=0)
        menuView_already = Menu(menuView, tearoff=0)
        # 创建三级菜单（第三条竖线排列）（标注设置_自动标注设置）
        menuLable_auto_model = Menu(menuLable_auto, tearoff=0)
        menuLable_auto_strategy = Menu(menuLable_auto, tearoff=0)
        """
        root.bind("<Control-n>", lambda event: self.newfile())
        root.bind("<Control-o>", lambda event: self.openfile())
        root.bind("<Control-s>", lambda event: self.savefile())
        root.bind("<Control-q>", lambda event: self.exit())
        """
        # 编辑菜单栏
        menubar.add_cascade(label="文件", menu=menuFile)
        menubar.add_cascade(label="标注设置", menu=menuLable)
        menubar.add_cascade(label="输出设置", menu=menuOutput)
        menubar.add_cascade(label="模式设置", menu=menuMethod)
        # menubar.add_cascade(label="视图", menu=menuView) # 待优化
        menubar.add_command(label="转换Excel为XML", command=self.trans_excxel_xml)
        menubar.add_command(label="智能预标注", command=self.OCR)

        # 编辑一级菜单（文件）
        menuFile.add_command(label="打开图片", command=self.openFile)
        menuFile.add_command(label="打开图片文件夹", command=self.openFileList)
        global filename
        global filedir
        menuFile.add_command(label="关闭图片（文件夹）", command=self.close_pic)
        # menuFile.add_separator()
        # menuFile.add_command(label="保存", accelerator="Ctrl+S", command=self.saveFile)
        menuFile.add_separator()
        menuFile.add_command(label="打开Excel文档", command=self.firstopenxlsFile)
        menuFile.add_command(label="关闭Excel文档", command=self.close_excel)
        menuFile.add_separator()
        menuFile.add_command(label="退出", command=self.master.destroy)
        # 编辑一级菜单（标注设置）
        menuLable.add_cascade(label="自动标注设置", menu=menuLable_auto)
        menuLable.add_cascade(label="手动标注设置（暂无）", menu=menuLable_manual)
        # 编辑一级菜单（输出设置）
        menuOutput.add_cascade(label='Excel', menu=menuOutput_xml)
        global var_picword
        var_picword = IntVar()
        menuOutput.add('checkbutton', label='图片+文字输出（默认选中）', variable=var_picword, command=self.print_)
        var_picword.set(1)
        gl.set_value('picword', True)
        # 编辑一级菜单（模式设置）
        global var_method
        var_method = StringVar()
        menuMethod.add('radiobutton', label='自动标注+手动修改（默认）', variable=var_method, value='auto_first', command=self.print_)
        menuMethod.add('radiobutton', label='手动标注', variable=var_method, value='manual', command=self.print_)
        var_method.set('auto_first')
        gl.set_value('method', 'auto_first')
        # 编辑一级菜单（视图）
        menuView.add_cascade(label="查看", menu=menuView_check)
        menuView.add_cascade(label="已标注图片", menu=menuView_already)
        # 编辑二级菜单（标注设置_自动标注设置）
        menuLable_auto.add_cascade(label='模型', menu=menuLable_auto_model)
        menuLable_auto.add_cascade(label='标注策略', menu=menuLable_auto_strategy)
        # 编辑二级菜单（输出设置_Excel）
        global var_Excel_num
        global var_Excel_first
        global var_Excel_auto
        var_Excel_num = StringVar()
        var_Excel_first = IntVar()
        var_Excel_auto = IntVar()
        # menuOutput_xml.add('radiobutton', label='单Excel输出（默认） （推荐，读入时会自动匹配）', variable=var_Excel_num, value='one_Excel', command=self.print_)
        # menuOutput_xml.add('radiobutton', label='逐张输出Excel （不推荐，再次读入时需手动匹配）', variable=var_Excel_num, value='Excels', command=self.print_)
        # menuOutput_xml.add_separator()
        menuOutput_xml.add('checkbutton', label='首行冻结（默认选中）', variable=var_Excel_first, command=self.print_)
        menuOutput_xml.add('checkbutton', label='自适应窗格大小（默认选中）', variable=var_Excel_auto, command=self.print_)
        var_Excel_num.set('one_Excel')
        var_Excel_first.set(1)
        gl.set_value('frozen', True)
        var_Excel_auto.set(1)
        gl.set_value('auto_width', True)
        # 编辑二级菜单（视图_查看）
        global var_View_look
        var_View_look = StringVar()
        menuView_check.add('radiobutton', label='图标（默认）', variable=var_View_look, value='icon', command=self.print_)
        menuView_check.add('radiobutton', label='列表', variable=var_View_look, value='list', command=self.print_)
        var_View_look.set('icon')
        # 编辑二级菜单（视图_已标注图片）
        global var_View_already
        var_View_already = StringVar()
        menuView_already.add('radiobutton', label='直接加载（默认）', variable=var_View_already, value='load', command=self.print_)
        menuView_already.add('radiobutton', label='逐张询问', variable=var_View_already, value='ask', command=self.print_)
        menuView_already.add('radiobutton', label='重新标记', variable=var_View_already, value='relabel', command=self.print_)
        var_View_already.set('load')
        # 编辑三级菜单（标注设置_自动标注设置_模型）
        global var_model
        var_model = StringVar()
        menuLable_auto_model.add('radiobutton', label='速度快（默认）', variable=var_model, value='speed_first', command=self.print_)
        menuLable_auto_model.add('radiobutton', label='精度高', variable=var_model, value='accuracy_first', command=self.print_)
        var_model.set('speed_first')
        gl.set_value('model', 'speed_first')
        # 编辑三级菜单（标注设置_自动标注设置_标注策略）
        global var_strategy
        var_strategy = StringVar()
        menuLable_auto_strategy.add('radiobutton', label='平衡模式（默认）', variable=var_strategy, value='balance', command=self.print_)
        menuLable_auto_strategy.add('radiobutton', label='广度优先', variable=var_strategy, value='wide_first', command=self.print_)
        menuLable_auto_strategy.add('radiobutton', label='准确度优先', variable=var_strategy, value='accuracy_first', command=self.print_)
        var_strategy.set('balance')
        gl.set_value('strategy', 'balance')
        self.master['menu'] = menubar

        """分割线"""
        self.line_sep = Canvas(self.master, width=2, height=Win_H-5, bg="black")
        self.line_sep.place(x=int(Win_W*2/3)-1, y=0)

        """打开Excel文件"""
        self.button_xls = Button(root, text="单击打开对应Excel文档", width=25, height=2, font=('黑体', '10'), command=self.firstopenxlsFile)
        self.button_xls.place(x=int(Win_W*5/6), y=int(Win_H/2), anchor='center')

        """打开（图片文件夹）"""
        self.button_xls = Button(root, text="打开（图片文件夹）", width=25, height=2, font=('黑体', '10'), command=self.openFileList)
        self.button_xls.place(x=int(Win_W*1/3), y=int(Win_H/2), anchor='center')

        """绑定窗口更新"""
        self.master.bind("<Configure>", self.Win_config)


    def print_(self):
        print("是否输出图像文字拼接：", "是" if var_picword.get() else "否", flush=True)
        print("Excel输出：", "单Excel输出" if var_Excel_num.get() == "one_Excel" else "逐张输出Excel", flush=True)
        print("\t首行冻结：", "打开" if var_Excel_first else "关闭", flush=True)
        if var_Excel_first:
            gl.set_value('frozen', True)
        else:
            gl.set_value('frozen', False)
        print("\t自适应窗格大小：", "打开" if var_Excel_auto else "关闭", flush=True)
        if var_method.get() == "auto_first":
            print("模式&标注设置：自动标注+手动修改", flush=True)
            print("\t模型：", "速度快" if var_model.get() == "speed_first" else "精度高", flush=True)
            print("\t标注策略：", "平衡模式" if var_strategy.get() == "balance" else \
                ("广度优先" if var_strategy.get() == "wide_first" else "准确度优先"), flush=True)
        else:
            print("模式设置：手动标注", flush=True)
        print("视图：", flush=True)
        print("\t查看：", "图标" if var_View_look.get() == "icon" else "列表", flush=True)
        print("\t已标注图像：", "直接加载" if var_View_already.get() == "load" else \
            ("逐张询问" if var_View_already.get() == "ask" else "重新标记"), flush=True)
        print("窗口信息：\n\t宽："+str(Win_W)+"\n\t高："+str(Win_H)+\
            "\n\tX坐标："+str(Win_X)+"\n\tY坐标："+str(Win_Y), flush=True)
        print("")
        if var_Excel_first.get(): # 设置首行冻结
            gl.set_value('frozen', True)
        else:
            gl.set_value('frozen', False)
        if var_Excel_auto.get(): # 设置自动宽度
            gl.set_value('auto_width', True)
        else:
            gl.set_value('auto_width', False)
        if var_picword.get():
            gl.set_value('picword', True)
        else:
            gl.set_value('picword', False)
        if var_method.get() == 'auto_first':
            gl.set_value('method', 'auto_first')
        else:
            gl.set_value('method', 'manual')
        if var_model.get() == 'speed_first':
            gl.set_value('model', 'speed_first')
        else:
            gl.set_value('model', 'accuracy_first')
        if var_strategy.get() == 'balance':
            gl.set_value('strategy', 'balance')
        elif var_strategy.get() == 'wide_first':
            gl.set_value('strategy', 'wide_first')
        else:
            gl.set_value('strategy', 'accuracy_first')

    """
       N
    W     E
       S
    """

    def openFile(self):
        print("进入openFile", flush=True)
        global mode
        mode = "file"
        global filename_bylist
        global filename
        global new_draw
        global prominent_rec
        try:
            self.saveExcel() # 尝试自动保存
        except:
            pass
        new_draw = []
        prominent_rec = []
        try:
            filename = askopenfile(filetypes=[("图片文件","*.jpeg"),\
                ("图片文件","*.jpg"),\
                ("图片文件","*.png")]).name
            # filename = "C:\\Users\\m1861\\Desktop\\SCD\\test_picture\\111111.png" # 便于调试
            print("打开文件：", filename, flush=True)
        except:
            pass
        self.showfile(filename)
        filename_bylist = []
        filename_bylist.append(filename) # 统一前后图按键，同时方便OCR
    
    def showfile(self, file):
        print("进入showfile", flush=True)
        global img_real
        img = cv2.imread(file)
        img_shape = img.shape
        img_real = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
        expect_W = int(Win_H*img_shape[1]/img_shape[0])
        expect_H = int(Win_H)
        global w, h
        global real_w, real_h
        real_w = img_shape[1]
        real_h = img_shape[0]
        size = (expect_W, expect_H) if expect_W < Win_W*2/3-3 else (int(Win_W*2/3)-3, int(Win_W*2/3*img_shape[0]/img_shape[1]))
        w = size[0]
        h = size[1]
        img = cv2.resize(img_real, size)
        global image
        image = ImageTk.PhotoImage(Image.fromarray(img.astype(np.uint8)))
        try:
            if self.imLabel:
                self.imLabel.destroy()
        except:
            pass
        self.imLabel=Label(self.master, image=image)
        global real_x, real_y, x_offset, y_offset
        real_x = (Win_W*2/3-size[0])/(2*Win_W)
        x_offset = -2
        real_y = (Win_H-size[1])/(2*Win_H)
        y_offset = 0
        self.imLabel.place(relx=real_x, x=x_offset, rely=real_y, y=y_offset, \
            width=int(min((Win_W * 2/3)- 3 -(real_x * Win_W + x_offset), w)))
        global CanMove
        CanMove=IntVar(self.master,value=0)
        global pic_x
        pic_x = IntVar(self.master,value=0)
        global pic_y
        pic_y = IntVar(self.master,value=0)
        '''
        global CanDraw
        CanDraw=IntVar(self.master,value=0)
        '''
        global rec_x
        rec_x = 0
        global rec_y
        rec_y = 0
        self.imLabel.bind("<Control-Button-1>", self.OnLeftButtonDown)
        self.imLabel.bind("<Control-ButtonRelease-1>", self.OnLeftButtonUp)
        self.imLabel.bind("<Control-B1-Motion>", self.OnLeftButtonMove)
        self.imLabel.bind("<Button-1>", self.BD_draw_rec) # 左键画框
        self.imLabel.bind("<ButtonRelease-1>", self.BU_draw_rec)
        self.imLabel.bind("<B1-Motion>", self.BM_draw_rec) # 实时更新正在画的方框
        self.imLabel.bind("<Button-3>", self.BD_del_rec) # 右键框选删除
        self.imLabel.bind("<ButtonRelease-3>", self.BU_del_rec)
        self.imLabel.bind("<B3-Motion>", self.BM_del_rec)
        self.imLabel.bind("<MouseWheel>", self.OnMouseWheel) # 滚轮放大缩小
        self.imLabel.bind("<Motion>", self.aux_line) # 鼠标进入画布，加辅助线
        self.imLabel.bind("<Leave>", self.mouse_leave) # 鼠标离开画布，去掉辅助线
        self.master.bind('<Down>', self.turn_right) # 顺时针旋转图像
        self.master.bind('<Up>', self.turn_left) # 逆时针旋转图像
        self.master.bind('<space>', self.rec_pic) # 恢复原大小
        self.openxlssheet()

    def rec_pic(self, event):
        try:
            self.saveExcel() # 尝试自动保存
        except:
            pass
        if mode == "file":
            self.showfile(filename) # 恢复原大小
        else:
            self.showfile(filename_bylist[pic_loc])

    def openFileList(self):
        global new_draw
        global prominent_rec
        try:
            self.saveExcel() # 尝试自动保存
        except:
            pass
        new_draw = []
        prominent_rec = []
        global filename
        filename = ""
        print("进入openFileList", flush=True)
        global mode
        mode = "filelist"
        global filedir
        filedir = askdirectory()
        print("打开目录：", filedir, flush=True)
        global filename_bylist
        global pic_loc
        pic_loc = 0
        filename_bylist = []
        pic_names = os.listdir(filedir) # 图片路径
        global pic_num
        pic_num = 0 # 用于统计文件数量
        for pic_name in pic_names:
            if pic_name[-4:] == ".jpg" or pic_name[-4:] == ".png":
                pic_num = pic_num + 1
                filename_bylist.append(filedir +"/"+ pic_name)
        print("pic_num:", pic_num, flush=True)
        # self.master.bind('<Right>', self.next_pic)
        self.master.bind('<Right>', self.next_pic)
        self.master.bind('<Left>', self.last_pic)
        # self.master.bind('<Left>', self.last_pic)
        global list_init
        list_init = 0
        self.showfile(filename_bylist[pic_loc])
    
    def next_pic(self, event):
        global pic_loc
        if mode == "file":
            messagebox.showwarning('警告','已到达最后一张照片')
            return
        pic_loc = (pic_loc + 1) if (pic_loc + 1) <= pic_num else (pic_loc)
        print("进入next_pic", flush=True)
        global new_draw
        global nwin
        try:
            nwin.destroy() # 尝试自动关闭修改窗口
        except:
            pass
        try:
            self.saveExcel() # 尝试自动保存
        except:
            pass
        new_draw = []
        global prominent_rec
        prominent_rec = []
        if pic_loc < pic_num:
            self.showfile(filename_bylist[pic_loc])
        else:
            pic_loc = pic_num - 1
            messagebox.showwarning('警告','已到达最后一张照片')
    
    def last_pic(self, event):
        global pic_loc
        if mode == "file":
            messagebox.showwarning('警告','已到达第一张照片')
            return
        pic_loc = (pic_loc - 1) if (pic_loc - 1) >= -1 else (pic_loc)
        print("进入last_pic", flush=True)
        global new_draw
        global nwin
        try:
            nwin.destroy() # 尝试自动关闭修改窗口
        except:
            pass
        try:
            self.saveExcel() # 尝试自动保存
        except:
            pass
        new_draw = []
        global prominent_rec
        prominent_rec = []
        if pic_loc >= 0:
            self.showfile(filename_bylist[pic_loc])
        else:
            pic_loc = 0
            messagebox.showwarning('警告','已到达第一张照片') # 融合单张照片情况

    def turn_right(self, event):
        global rectan, new_draw, xlarr
        print("进入rot_pic", flush=True)
        global img_real
        img_real = np.rot90(img_real)
        img_real = np.rot90(img_real)
        img_real = np.rot90(img_real)
        choose = str(messagebox.askyesnocancel('顺时针旋转图片', '是否删除此图片已标记框和数据？\n“取消”放弃旋转'))
        print("choose:", choose, flush=True)
        if choose == "True":
            rectan = ["text_box_position"]
            new_draw = []
            xlarr = ["text"]
            print("rectan:", rectan, "new_draw:", new_draw, "xlarr:", xlarr, flush=True)
            self.saveExcel()
        elif choose == "None":
            img_real = np.rot90(img_real)
        img_save = cv2.cvtColor(img_real, cv2.COLOR_BGR2RGB)
        if mode == "file":
            cv2.imwrite(filename, img_save)
        else:
            cv2.imwrite(filename_bylist[pic_loc], img_save)
        if mode == "file":
            self.showfile(filename) # 重新打开
        else:
            self.showfile(filename_bylist[pic_loc])

    def turn_left(self, event):
        global rectan, new_draw, xlarr
        print("进入rot_pic", flush=True)
        global img_real
        img_real = np.rot90(img_real)
        choose = str(messagebox.askyesnocancel('逆时针旋转图片', '是否删除此图片已标记框和数据？\n“取消”放弃旋转'))
        print("choose:", choose, flush=True)
        if choose == "True":
            rectan = ["text_box_position"]
            new_draw = []
            xlarr = ["text"]
            print("rectan:", rectan, "new_draw:", new_draw, "xlarr:", xlarr, flush=True)
            self.saveExcel()
        elif choose == "None":
            img_real = np.rot90(img_real)
            img_real = np.rot90(img_real)
            img_real = np.rot90(img_real)
        img_save = cv2.cvtColor(img_real, cv2.COLOR_BGR2RGB)
        if mode == "file":
            cv2.imwrite(filename, img_save)
        else:
            cv2.imwrite(filename_bylist[pic_loc], img_save)
        if mode == "file":
            self.showfile(filename) # 重新打开
        else:
            self.showfile(filename_bylist[pic_loc])

    def close_pic(self):
        print("进入close_pic", flush=True)
        try:
            self.saveExcel()
        except:
            pass
        global filename
        global filedir
        global filename_bylist
        filename=""
        filedir=""
        filename_bylist = []
        try:
            self.imLabel.destroy()
        except:
            pass
        if excelname != "":
            messagebox.showwarning('警告','请您首先打开图片文件！')

    def aux_line(self, event):
        img_real_temp = cv2.resize(img_real, (w,h)) # 限制线宽，保证观感更好
        try:
            for rec in rectan[1:]:
                # print(rec, flush=True)
                cv2.rectangle(img_real_temp, (int(rec[0][0]*w/real_w), int(rec[0][1]*h/real_h)), \
                    (int(rec[2][0]*w/real_w), int(rec[2][1]*h/real_h)), (255,0,0), 2)
        except:
            pass
        try:
            for content in new_draw:
                # print(content, flush=True)
                cv2.rectangle(img_real_temp, (int(content[0][0]*w/real_w), int(content[0][1]*h/real_h)), \
                    (int(content[2][0]*w/real_w), int(content[2][1]*h/real_h)), (255,0,0), 2)
        except:
            pass # Excel中无框的信息
        ptStart = (event.x, 0)
        ptEnd = (event.x, h)
        point_color = (0, 255, 0) # 绿色
        thickness = 1 # 最细
        lineType = 1
        cv2.line(img_real_temp, ptStart, ptEnd, point_color, thickness, lineType)
        ptStart = (0, event.y)
        ptEnd = (w, event.y)
        cv2.line(img_real_temp, ptStart, ptEnd, point_color, thickness, lineType)
        img = cv2.resize(img_real_temp, (w,h))
        img = img[:, 0:int((Win_W * 2/3)- 3 -(real_x * Win_W + x_offset))]
        global image_tk
        image_tk = ImageTk.PhotoImage(Image.fromarray(img.astype(np.uint8)))
        self.imLabel['image'] = image_tk
        self.imLabel.image = image_tk
        self.imLabel.place(relx=real_x, x=x_offset, rely=real_y, y=y_offset, \
            width=int(min((Win_W * 2/3)- 3 -(real_x * Win_W + x_offset), w)))
        self.imLabel.update()
        
    def mouse_leave(self, event):
        # print("进入mouse_leave", flush=True)
        img_real_temp = cv2.resize(img_real, (w,h)) # 限制线宽，保证观感更好
        try:
            for rec in rectan[1:]:
                # print(rec, flush=True)
                cv2.rectangle(img_real_temp, (int(rec[0][0]*w/real_w), int(rec[0][1]*h/real_h)), \
                    (int(rec[2][0]*w/real_w), int(rec[2][1]*h/real_h)), (255,0,0), 2)
        except:
            pass
        try:
            for content in new_draw:
                # print(content, flush=True)
                cv2.rectangle(img_real_temp, (int(content[0][0]*w/real_w), int(content[0][1]*h/real_h)), \
                    (int(content[2][0]*w/real_w), int(content[2][1]*h/real_h)), (255,0,0), 2)
        except:
            pass # Excel中无框的信息
        img = cv2.resize(img_real_temp, (w,h))
        img = img[:, 0:int((Win_W * 2/3)- 3 -(real_x * Win_W + x_offset))]
        global image_tk
        image_tk = ImageTk.PhotoImage(Image.fromarray(img.astype(np.uint8)))
        self.imLabel['image'] = image_tk
        self.imLabel.image = image_tk
        self.imLabel.place(relx=real_x, x=x_offset, rely=real_y, y=y_offset, \
            width=int(min((Win_W * 2/3)- 3 -(real_x * Win_W + x_offset), w)))
        self.imLabel.update()


    def firstopenxlsFile(self):
        print("进入firstopenxlsFile", flush=True)
        global excelname
        try:
            self.saveExcel() # 尝试自动保存
        except:
            pass
        excelname = askopenfile(filetypes=[("Excel文件","*.xlsx")]).name
        # excelname = "C:\\Users\\m1861\\Desktop\\SCD\\test_picture\\test.xls" # 便于调试
        self.openxlssheet()

    def openxlssheet(self):
        print("进入openxlssheet", flush=True)
        try:
            print("打开Excel文件：", excelname, flush=True)
        except:
            messagebox.showwarning('警告','请您首先打开Excel文件后再继续操作！')
        global excel_w
        excel_w = load_workbook(excelname)
        global sheet_names
        global file_name_real
        sheet_names = excel_w.get_sheet_names()
        global xlarr
        global rectan
        global new_draw
        new_draw = []
        try:
            if mode == "file":
                file_name_real = os.path.split(filename)
                for i, name in enumerate(sheet_names):
                    if name == file_name_real[1]:
                        sheet = excel_w[file_name_real[1]]
                        break
                    else:
                        if i == len(sheet_names)-1 and filename != "": # 新增关闭操作后打的补丁...
                            messagebox.showwarning('警告','无对应数据，已新建sheet！')
                            xlarr = ['text']
                            rectan = ['text_box_position']
                            self.make_tree()
                            return
            else:
                file_name_real = os.path.split(filename_bylist[pic_loc])
                for i, name in enumerate(sheet_names):
                    if name == file_name_real[1]:
                        sheet = excel_w[file_name_real[1]]
                        break
                    else:
                        if i == len(sheet_names)-1 and filename_bylist != []: # 新增关闭操作后打的补丁...
                            messagebox.showwarning('警告','无对应数据，已新建sheet！')
                            xlarr = ['text']
                            rectan = ['text_box_position']
                            self.make_tree()
                            return
        except:
            messagebox.showwarning('警告','请您首先打开图片文件！')
            return
        global nrows, ncols
        nrows = sheet.max_row
        ncols = sheet.max_column - 1 # 框的位置和置信度不需要显示出来
        xlarr = []
        rectan = []
        for r in range(nrows):
            arr = []
            rectan_temp = []
            for c in range(ncols):
                arr.append(sheet.cell(row=r+1, column=c+1).value)
                rectan_temp.append(sheet.cell(row=r+1, column=c+2).value)
            xlarr.append(arr[0])
            rectan.append(rectan_temp[0]) # 交给画框程序去画框
        rectan_format = []
        rectan_format.append(rectan[0])
        for rec in rectan[1:]:
            print(rec, flush=True)
            rectan_format.append(literal_eval(rec))
        rectan = rectan_format
        self.make_tree()

    def make_tree(self): # 分离显示程序，更方便编辑
        print("进入make_tree", flush=True)
        global tree
        global tree2
        global list_init
        try:
            tree.destroy()
        except:
            pass
        colname = []
        for c in range(1):
            colname.append(c)
        tree = ttk.Treeview(self.master, show='headings', columns=colname, selectmode = 'browse') # 单行选中模式
        for c in range(1):
            tree.column(c, anchor='center')
            nametofont("TkHeadingFont").configure(family="Times New Roman", size=18, weight="bold")
            tree.heading(c, text="标注の内容") # 显示标题
        global titles
        titles = [xlarr[0], rectan[0]]
        print("title:", titles)
        # del(xlarr[0]) # 删除标题行
        # del(rectan[0]) # 删除标题行
        for i in range(len(xlarr)-1): # 除去首行
            tree.insert('', i, values=xlarr[i+1]) # 显示内容
        if mode == "file":
            tree.place(x=(int(Win_W)+int((Win_W*2/3)))/2+3, y=0, anchor='n', width=int(Win_W)-int((Win_W*2/3)+6), height=Win_H)
            try:
                tree2.destroy()
            except:
                pass
        elif mode == "filelist":
            tree.place(x=(int(Win_W)+int((Win_W*2/3)))/2+3, y=0, anchor='n', width=int(Win_W)-int((Win_W*2/3)+6), height=Win_H-200)
        print("tree_build")
        tree.bind('<Double-Button-1>', self.edit) # 监控鼠标双击
        tree.bind('<Delete>', self.delete_kuanger) # 监控delete按键
        # btn1 = Button(self.master, text='保存', command=self.saveExcel)
        # btn1.place(x=(int(Win_W)+int((Win_W*2/3)))/2+3, y=Win_H-30, anchor='center', width=80, height=40)
        if mode == "filelist" and list_init == 0:
            list_init = 1
            tree2 = ttk.Treeview(self.master, show='headings', columns=colname, selectmode = 'browse') # 单行选中模式
            for c in range(1):
                tree2.column(c, anchor='center')
                nametofont("TkHeadingFont").configure(family="Times New Roman", size=15, weight="bold")
                tree2.heading(c, text="文件夹中の文件列表") # 显示标题
            for i in range(len(filename_bylist)):
                tree2.insert('', i, values=os.path.split(filename_bylist[i])[1]) # 显示内容
            tree2.place(x=(int(Win_W)+int((Win_W*2/3)))/2+3, y=Win_H-194, anchor='n', width=int(Win_W)-int((Win_W*2/3)+6), height=194)
            print("tree2_build")
            tree2.bind('<Double-Button-1>', self.changepic) # 监控鼠标双击
        self.draw_rec()

    def edit(self, event):
        print("进入edit", flush=True)
        global nwin
        global tree
        global enty
        global colint
        global sitem
        try:
            nwin.destroy()
        except:
            pass
        for item in tree.selection():
            ttext = tree.item(item, 'values')
            sitem = item
        global prominent_rec
        all_rec = []
        try:
            all_rec = rectan + new_draw
        except:
            all_rec = rectan
        prominent_rec = all_rec[int(str(sitem[1:]), 16)]
        self.draw_rec()
        col = tree.identify_column(event.x)
        colint = int(str(col.replace('#', '')))
        nwin = Tk() # 子窗口
        nwin.title("修改框内文字") # 子窗口名称
        nwin.resizable(False, False) # 规定窗口不可缩放
        '''
        win_test = Tk()
        win_test.winfo_screenheight()
        loc_x = (win_test.winfo_screenwidth()-400) if (event.x_root-200+400) > win_test.winfo_screenwidth() else (event.x_root-200+400)
        loc_y = (win_test.winfo_screenheight()-60) if (event.x_root+30+60) > win_test.winfo_screenheight() else (event.x_root-30+60)
        print(loc_x, loc_y)
        print(win_test.winfo_screenwidth(), win_test.winfo_screenheight())
        win_test.destroy()
        '''
        loc_x = int(event.x_root)-200
        loc_y = int(event.y_root)+30
        nwin.geometry("400x60+"+str(loc_x)+"+"+str(loc_y)) # 使弹出窗口有较好的用户体验
        label1 = Label(nwin, text="修改:")
        label1.place(relx=0, x=5, rely=0.5, anchor="w", width=30, height=25)
        enty = Text(nwin, wrap=WORD)
        enty = Entry(nwin)
        enty.place(relx=0, x=40, rely=0.5, anchor="w", width=320, height=25)
        btn = Label(nwin, text='确认')
        btn.place(relx=0, x=365, rely=0.5, anchor="w", width=30, height=25)
        btn.bind('<Button-1>', self.getv)
        self.master.bind('<Return>', self.getv)
        enty.bind('<Return>', self.getv)
        enty.focus()
        enty.insert('end', ttext[colint-1]) # 编辑框显示值
        nwin.mainloop()

    def getv(self, event):
        print("进入getv", flush=True)
        global nwin
        global enty
        global tree
        global sitem
        global colint
        global xlarr # 统一在这里更新文字数据
        editxt = enty.get()
        tree.set(sitem, (colint-1), editxt)
        print("sitem:", int(str(sitem[1:]), 16))
        print("colint-1:", colint-1, flush=True)
        print("editxt:", editxt, flush=True)
        # print("xlarr:", xlarr, flush=True)
        # print("len(xlarr):", len(xlarr), flush=True)
        xlarr[int(str(sitem[1:]), 16)] = editxt
        nwin.destroy()

    def changepic(self, event):
        global tree2
        global pic_loc
        for item in tree2.selection():
            ttext = tree2.item(item, 'values')
            sitem = item
        print("ttext", ttext[0], flush=True)
        for i, pic in enumerate(filename_bylist):
            if ttext[0] == os.path.split(pic)[1]:
                pic_loc = i
                break
        self.showfile(pic)


    def saveExcel(self):
        print("进入saveExcel", flush=True)
        global excel_w
        global tree
        global titles
        global excelname
        global sheet_names
        all_rec = []
        all_rec = rectan + new_draw
        # print("1", flush=True)
        print("save_sheet_name:", file_name_real[1],flush=True)
        for i, excel_pic_name in enumerate(sheet_names): # 防止其他信息丢失，一律重写
            if excel_pic_name == file_name_real[1]:
                sheet_wt = excel_w[file_name_real[1]]
                break
            if i == len(sheet_names)-1:
                print("新建sheet", flush=True)
                sheet_wt = excel_w.create_sheet()
                sheet_wt.title = file_name_real[1] # 若没有则新建一个（考虑写到菜单栏里？）
                sheet_names += file_name_real[1] # 否则下次还会新建！
                break
        # print("2", flush=True)
        j = 0
        for i in range(max(sheet_wt.max_row, len(xlarr))): # 取二者较大值
            j = j + 1
            if i < len(xlarr):
                sheet_wt.cell(row=j, column=1).value=str(xlarr[i]) # 转换为字符串，提取value中的内容
                sheet_wt.cell(row=j, column=2).value=str(all_rec[i]) # 保存对应的框
            else:
                sheet_wt.delete_rows(j)
                j = j - 1 # 删掉一行之后下一行变成当前行；sheet_wt.max_row会自动减少
        # print("3", flush=True)
        len_byte_text = 5
        len_byte_text_box_position = 5
        for words in xlarr:
            len_byte_text = max(len_byte(str(words)), len_byte_text)
            # print("len_byte(str(words)):", len_byte(str(words)),flush=True)
        for words in all_rec:
            len_byte_text_box_position = max(len_byte(str(words)), len_byte_text)
            # print("len_byte(str(words)):", len_byte(str(words)),flush=True)
        # print("4", flush=True)
        if(gl.get_value('auto_width')==True): # 设置自动宽度
            sheet_wt.column_dimensions["A"].width = (len_byte_text + 1)
            sheet_wt.column_dimensions["B"].width = (len_byte_text_box_position + 1)
        else:
            pass
        if(gl.get_value('frozen')==True): # 设置首行冻结
            sheet_wt.freeze_panes = 'A2'
        else:
            sheet_wt.freeze_panes = 'A1'
        # print("5", flush=True)
        excel_w.save(excelname)
        # print("6", flush=True)
        # messagebox.showinfo('提示', '保存成功') # 已添加自动保存机制

    def close_excel(self):
        print("进入close_excel", flush=True)
        global tree
        try:
            self.saveExcel()
        except:
            pass
        global excelname
        global xlarr
        global rectan
        global tree2
        excelname = ""
        xlarr = []
        rectan = []
        tree.destroy()
        try:
            tree2.destroy()
        except:
            pass

    def new_kuanger(self):
        print("进入new_kuanger", flush=True)
        global xlarr
        xlarr.append("********NULL********") # 新建文字数据
        global nwin
        global enty2
        try:
            nwin.destroy()
        except:
            pass
        nwin = Tk() # 子窗口
        nwin.title("修改框内文字") # 子窗口名称
        nwin.resizable(False, False) # 规定窗口不可缩放
        loc_x, loc_y = pag.position() # 获取当前鼠标绝对位置
        nwin.geometry("400x60+"+str(loc_x)+"+"+str(loc_y)) # 使弹出窗口有较好的用户体验
        nwin.protocol("WM_DELETE_WINDOW", self.close_input_win)
        label1 = Label(nwin, text="修改:")
        label1.place(relx=0, x=5, rely=0.5, anchor="w", width=30, height=25)
        enty2 = Text(nwin, wrap=WORD)
        enty2 = Entry(nwin)
        enty2.place(relx=0, x=40, rely=0.5, anchor="w", width=320, height=25)
        btn = Label(nwin, text='确认')
        btn.place(relx=0, x=365, rely=0.5, anchor="w", width=30, height=25)
        btn.bind('<Button-1>', self.new_value)
        self.master.bind('<Return>', self.new_value)
        enty2.bind('<Return>', self.new_value)
        nwin.mainloop()

    def close_input_win(self): # 防止意外退出，禁止手动关闭此窗口（可通过"enter"键关闭）
        pass

    def new_value(self, event):
        print("进入new_value", flush=True)
        global xlarr
        editxt = enty2.get()
        if editxt == "":
            xlarr[-1] = "********NULL********" # 向用户展示有这个框
        else:
            xlarr[-1] = editxt
        nwin.destroy()
        self.make_tree()
    
    def delete_kuanger(self, event):
        print("进入delete_kuanger", flush=True)
        global xlarr
        global rectan
        global new_draw
        for item in tree.selection():
            sitem = item
        del xlarr[int(str(sitem[1:]), 16)]
        if int(str(sitem[1:]), 16) < len(rectan):
            del rectan[int(str(sitem[1:]), 16)]
        else:
            del new_draw[int(str(sitem[1:]), 16)-len(rectan)]
        self.make_tree()


    def draw_rec(self):
        print("进入draw_rec", flush=True)
        img_real_temp = cv2.resize(img_real, (w,h)) # 限制线宽，保证观感更好
        try:
            for rec in rectan[1:]:
                # print(rec, flush=True)
                cv2.rectangle(img_real_temp, (int(rec[0][0]*w/real_w), int(rec[0][1]*h/real_h)), \
                    (int(rec[2][0]*w/real_w), int(rec[2][1]*h/real_h)), (255,0,0), 2)
        except:
            pass
        try:
            for content in new_draw:
                # print(content, flush=True)
                cv2.rectangle(img_real_temp, (int(content[0][0]*w/real_w), int(content[0][1]*h/real_h)), \
                    (int(content[2][0]*w/real_w), int(content[2][1]*h/real_h)), (255,0,0), 2)
        except:
            pass # Excel中无框的信息
        try:
            cv2.rectangle(img_real_temp, (int(prominent_rec[0][0]*w/real_w), int(prominent_rec[0][1]*h/real_h)), \
                (int(prominent_rec[2][0]*w/real_w), int(prominent_rec[2][1]*h/real_h)), (0,0,255), 2)
        except:
            pass
        print("", flush=True)
        img = img_real_temp[:, 0:int((Win_W * 2/3)- 3 -(real_x * Win_W + x_offset))]
        global image_tk
        image_tk = ImageTk.PhotoImage(Image.fromarray(img.astype(np.uint8)))
        self.imLabel['image'] = image_tk
        self.imLabel.image = image_tk

    def draw_rec_temp(self):
        # print("进入temp画框程序", flush=True)
        # img_real_temp = copycopy.copy(img_real) # 必须要这样复制！否则只是复制了“指针”
        img_real_temp = cv2.resize(img_real, (w,h)) # 限制线宽，保证观感更好
        try:
            for content2 in new_draw_temp:
                # print(content2, flush=True)
                cv2.rectangle(img_real_temp, (int(content2[0][0]*w/real_w), int(content2[0][1]*h/real_h)), \
                    (int(content2[2][0]*w/real_w), int(content2[2][1]*h/real_h)), (255,0,0), 2)
        except:
            pass # 鼠标不处于“move”状态
        # print("", flush=True)
        img = img_real_temp[:, 0:int((Win_W * 2/3)- 3 -(real_x * Win_W + x_offset))]
        global image_tk
        image_tk = ImageTk.PhotoImage(Image.fromarray(img.astype(np.uint8)))
        self.imLabel['image'] = image_tk
        self.imLabel.image = image_tk

    def del_rec_temp(self):
        img_real_temp = cv2.resize(img_real, (w,h)) # 限制线宽，保证观感更好
        try:
            for content2 in new_del_temp:
                # print(content2, flush=True)
                cv2.rectangle(img_real_temp, (int(content2[0][0]*w/real_w), int(content2[0][1]*h/real_h)), \
                    (int(content2[2][0]*w/real_w), int(content2[2][1]*h/real_h)), (0,0,255), 2)
        except:
            pass # 鼠标不处于“move”状态
        try:
            for rec in rectan[1:]:
                # print(rec, flush=True)
                cv2.rectangle(img_real_temp, (int(rec[0][0]*w/real_w), int(rec[0][1]*h/real_h)), \
                    (int(rec[2][0]*w/real_w), int(rec[2][1]*h/real_h)), (255,0,0), 2)
        except:
            pass
        try:
            for content in new_draw:
                # print(content, flush=True)
                cv2.rectangle(img_real_temp, (int(content[0][0]*w/real_w), int(content[0][1]*h/real_h)), \
                    (int(content[2][0]*w/real_w), int(content[2][1]*h/real_h)), (255,0,0), 2)
        except:
            pass # Excel中无框的信息
        # print("", flush=True)
        img = img_real_temp[:, 0:int((Win_W * 2/3)- 3 -(real_x * Win_W + x_offset))]
        global image_tk
        image_tk = ImageTk.PhotoImage(Image.fromarray(img.astype(np.uint8)))
        self.imLabel['image'] = image_tk
        self.imLabel.image = image_tk

    def OnLeftButtonDown(self, event):
        global CanMove
        pic_x.set(event.x)
        pic_y.set(event.y)
        CanMove.set(1)

    def OnLeftButtonUp(self, event):
        global CanMove
        CanMove.set(0)

    def OnLeftButtonMove(self, event):
        global CanMove
        global real_x, real_y, x_offset, y_offset
        global w, h, img_real
        if CanMove.get()==0:
            return
        # 加框（由于加框函数不对原图做修改，只能在这里单独加框了...）
        img_real_temp = cv2.resize(img_real, (w,h)) # 限制线宽，保证观感更好
        try:
            for rec in rectan[1:]:
                print(rec, flush=True)
                cv2.rectangle(img_real_temp, (int(rec[0][0]*w/real_w), int(rec[0][1]*h/real_h)), \
                    (int(rec[2][0]*w/real_w), int(rec[2][1]*h/real_h)), (255,0,0), 2)
        except:
            pass
        try:
            for content in new_draw:
                print(content, flush=True)
                cv2.rectangle(img_real_temp, (int(content[0][0]*w/real_w), int(content[0][1]*h/real_h)), \
                    (int(content[2][0]*w/real_w), int(content[2][1]*h/real_h)), (255,0,0), 2)
        except:
            pass # Excel中无框的信息
        newx=(event.x-pic_x.get())
        newy=(event.y-pic_y.get())
        x_offset += newx
        y_offset += newy
        if w > (Win_W * 2/3)- 3 -(real_x * Win_W + x_offset):
            # 渲染
            img = cv2.resize(img_real_temp, (w,h))
            img = img[:, 0:int((Win_W * 2/3)- 3 -(real_x * Win_W + x_offset))]
            global image_tk
            image_tk = ImageTk.PhotoImage(Image.fromarray(img.astype(np.uint8)))
            self.imLabel['image'] = image_tk
            self.imLabel.image = image_tk
        self.imLabel.place(relx=real_x, x=x_offset, rely=real_y, y=y_offset, \
            width=int(min((Win_W * 2/3)- 3 -(real_x * Win_W + x_offset), w)))
        self.imLabel.update()


    def BD_draw_rec(self, event):
        global rec_x, rec_y
        global last_state
        last_state = 0
        rec_x = event.x
        rec_y = event.y
        # print("rec_x:{} real_rec_x:{}".format(rec_x, rec_x*real_w/w), flush=True)
        # print("real_w:{} w:{}".format(real_w, w), flush=True)
        # print("rec_y:{} real_rec_y:{}".format(rec_y, rec_y*real_h/h), flush=True)
        # print("real_h:{} h:{}".format(real_h, h), flush=True)

    def BU_draw_rec(self, event):
        global new_draw
        for enter_bug_solve in new_draw: # 鼠标点击确认的一个bug的解决...具体原因尚未查清
            if [enter_bug_solve[0][0], enter_bug_solve[0][1]] == [int(rec_x*real_w/w), int(rec_y*real_h/h)]:
                return
        try:
            new_draw.append([[int(rec_x*real_w/w), int(rec_y*real_h/h)], [int(event.x*real_w/w), int(rec_y*real_h/h)], \
                [int(event.x*real_w/w), int(event.y*real_h/h)], [int(rec_x*real_w/w), int(event.y*real_h/h)]])
            self.new_kuanger()
        except:
            new_draw = []
            messagebox.showwarning('警告','请您首先打开Excel文件后再画框！')
            return

    def BM_draw_rec(self, event):
        global last_state
        global new_draw_temp
        new_draw_temp = []
        # new_draw_temp.clear()
        new_draw_temp.append([[int(rec_x*real_w/w), int(rec_y*real_h/h)], [int(event.x*real_w/w), int(rec_y*real_h/h)], \
            [int(event.x*real_w/w), int(event.y*real_h/h)], [int(rec_x*real_w/w), int(event.y*real_h/h)]])
        self.draw_rec_temp()

    def BD_del_rec(self, event):
        global del_x, del_y
        global last_state
        last_state = 0
        del_x = event.x
        del_y = event.y

    def BU_del_rec(self, event):
        try:
            all_rec = rectan[1:] + new_draw
        except:
            messagebox.showwarning('警告','请您首先打开Excel文件后再删除区域框！')
            return
        len_rectan = len(rectan[1:]) # 除去第一位！！！
        print("len_rectan:", len_rectan, flush=True)
        print("len_all_rec:", len(all_rec), flush=True)
        dellist = []
        print("开始减法！！！")
        minx2 = min(int(del_x*real_w/w), int(event.x*real_w/w))
        miny2 = min(int(del_y*real_h/h), int(event.y*real_h/h))
        maxx2 = max(int(del_x*real_w/w), int(event.x*real_w/w))
        maxy2 = max(int(del_y*real_h/h), int(event.y*real_h/h))
        print("minx2,miny2,maxx2,maxy2:", minx2,miny2,maxx2,maxy2, flush=True)
        for i, rec in enumerate(all_rec): # 需要额外添加考虑被包围的情况！
            minx1 = min(rec[0][0], rec[2][0])
            miny1 = min(rec[0][1], rec[2][1])
            maxx1 = max(rec[2][0], rec[0][0])
            maxy1 = max(rec[2][1], rec[0][1])
            print("minx1,miny1,maxx1,maxy1:", minx1,miny1,maxx1,maxy1, flush=True)
            minx = max(minx1, minx2)
            miny = max(miny1, miny2)
            maxx = min(maxx1, maxx2)
            maxy = min(maxy1, maxy2)
            if minx2 > minx1 and maxx2 < maxx1 and miny2 > miny1 and maxy2 < maxy1:
                continue
            else:
                if minx <= maxx and miny <= maxy: # 不写等号删不了“点”
                    dellist.append([i, rec])
        del_n = len(dellist)
        if del_n != 0:
            choose = str(messagebox.askyesno('删除', '是否删除被框选的'+str(del_n)+'个标注框？'))
            if choose == "True":
                print("dellist:", dellist, flush=True)
                for i in range(len(dellist)):
                    j = len(dellist) - 1 - i # 倒序删除防止乱序
                    print("dellist[j][0]:", dellist[j][0], flush=True)
                    print("dellist[j][1]:", dellist[j][1], flush=True)
                    if dellist[j][0] < len_rectan: # 序号！不是j！
                        rectan.remove(dellist[j][1])
                        print("删除文字：", xlarr[int(dellist[j][0])+1], flush=True)
                        del xlarr[int(dellist[j][0])+1]
                    else:
                        new_draw.remove(dellist[j][1])
                        del xlarr[int(dellist[j][0])+1]
            else:
                pass
        self.make_tree()

    def BM_del_rec(self, event):
        global last_state
        global new_del_temp
        new_del_temp = []
        new_del_temp.append([[int(del_x*real_w/w), int(del_y*real_h/h)], [int(event.x*real_w/w), int(del_y*real_h/h)], \
            [int(event.x*real_w/w), int(event.y*real_h/h)], [int(del_x*real_w/w), int(event.y*real_h/h)]])
        self.del_rec_temp()

    def OnMouseWheel(self, event):
        global w, h, img_real
        global real_x, real_y, x_offset, y_offset
        # event.x是鼠标当前位置相对于图片边缘的x坐标
        # real_x * Win_W + x_offset是当前图片左上角的x坐标
        # event.x + (real_x * Win_W + x_offset)为鼠标相对窗口左上角的位置
        x = root.winfo_x()
        y = root.winfo_y()
        # print("event.x:", event.x, flush=True)
        # print("鼠标相对窗口左上角的位置：", event.x + (real_x * Win_W + x_offset), flush=True)
        # print("")
        # 加框（由于加框函数不对原图做修改，只能在这里单独加框了...）
        img_real_temp = cv2.resize(img_real, (w,h)) # 限制线宽，保证观感更好
        try:
            for rec in rectan[1:]:
                print(rec, flush=True)
                cv2.rectangle(img_real_temp, (int(rec[0][0]*w/real_w), int(rec[0][1]*h/real_h)), \
                    (int(rec[2][0]*w/real_w), int(rec[2][1]*h/real_h)), (255,0,0), 2)
        except:
            pass
        try:
            for content in new_draw:
                print(content, flush=True)
                cv2.rectangle(img_real_temp, (int(content[0][0]*w/real_w), int(content[0][1]*h/real_h)), \
                    (int(content[2][0]*w/real_w), int(content[2][1]*h/real_h)), (255,0,0), 2)
        except:
            pass # Excel中无框的信息
        if event.delta>0:
            w = int(w * 1.1)
            h = int(h * 1.1)
            x_offset = x_offset - 0.1 * event.x
            y_offset = y_offset - 0.1 * event.y
        else:
            w = int(w * 0.90909)
            h = int(h * 0.90909)
            x_offset = x_offset + (1 - 0.90909) * event.x
            y_offset = y_offset + (1 - 0.90909) * event.y
        # 渲染
        img = cv2.resize(img_real_temp, (w,h))
        if w > (Win_W * 2/3)- 3 -(real_x * Win_W + x_offset):
            img = img[:, 0:int((Win_W * 2/3)- 3 -(real_x * Win_W + x_offset))]
        global image_tk
        image_tk = ImageTk.PhotoImage(Image.fromarray(img.astype(np.uint8)))
        # image_tk = ImageTk.PhotoImage(image.resize((w,h)))
        self.imLabel['image'] = image_tk
        self.imLabel.image = image_tk
        self.imLabel.place(relx=real_x, x=x_offset, rely=real_y, y=y_offset, \
            width=min((Win_W * 2/3)- 3 -(real_x * Win_W + x_offset), w))
        self.imLabel.update()
        # self.master.update()


    def trans_excxel_xml(self):
        try:
            print(sheet_names)
        except:
            messagebox.showwarning('警告','请您首先打开Excel文件')
            return
        try:
            self.saveExcel() # 尝试自动保存
        except:
            pass
        for name in sheet_names:
            sheet = excel_w[name]
            nrows = sheet.max_row
            ncols = sheet.max_column - 1
            xlarr = []
            rectan = []
            for r in range(nrows):
                arr = []
                rectan_temp = []
                for c in range(ncols):
                    arr.append(sheet.cell(row=r+1, column=c+1).value)
                    rectan_temp.append(sheet.cell(row=r+1, column=c+2).value)
                xlarr.append(arr[0])
                rectan.append(rectan_temp[0])
            rectan_format = []
            rectan_format.append(rectan[0])
            for rec in rectan[1:]:
                print(rec, flush=True)
                rectan_format.append(literal_eval(rec))
            rectan = rectan_format
            print("rectan:", rectan, flush=True)
            print("xlarr:", xlarr, flush=True)
            root = etree.Element('annotation')
            if len(xlarr) > 0 and len(rectan) == len(xlarr):
                child1 = []
                child2 = []
                element = ['x', 'y', 'width', 'height', 'content']
                for i in range(len(xlarr)):
                    if i == 0:
                        continue
                    child1.append(etree.SubElement(root, 'project'))
                    x = min(rectan[i][0][0], rectan[i][1][0], rectan[i][2][0], rectan[i][3][0])
                    y = min(rectan[i][0][1], rectan[i][1][1], rectan[i][2][1], rectan[i][3][1])
                    width = abs(rectan[i][2][0] - rectan[i][0][0]) # 保存的时候有可能是反的，需要在此处改成正的！
                    height = abs(rectan[i][2][1] - rectan[i][0][1]) # 同上
                    content = xlarr[i]
                    for j in range(5):
                        child2.append(etree.SubElement(child1[i-1], element[j]))
                    child2[0].text = str(x)
                    child2[1].text = str(y)
                    child2[2].text = str(width)
                    child2[3].text = str(height)
                    child2[4].text = str(content)
                    child2 = []
                tree = etree.ElementTree(root)
                try:
                    os.remove('C:\\Users\\m1861\\Desktop\\SCD\\XML\\'+name[:-4]+'.xml')
                except:
                    pass
                tree.write('C:\\Users\\m1861\\Desktop\\SCD\\XML\\'+name[:-4]+'.xml', pretty_print=True, xml_declaration=True, encoding='utf-8')


    def Win_config(self, event):
        global Win_W
        global Win_H
        global Win_X
        global Win_Y
        if Win_W == event.width and Win_H == event.height:
            Win_X = event.x
            # print("Win_X:", Win_X, flush=True)
            Win_Y = event.y
            # print("Win_Y:", Win_Y, flush=True)

    
    def OCR(self):
        if gl.get_value('method')=='manual':
            messagebox.showwarning('警告','您选择的是手动标注！')
            return
        try:
            if len(filename_bylist) == 0:
                messagebox.showwarning('警告','请您首先打开图片文件（夹）！')
                return
        except:
            messagebox.showwarning('警告','请您首先打开图片文件（夹）！')
            return
        choose = str(messagebox.askyesno('请问...', '是否按照您的配置进行自动标注推算？'))
        if choose == "True":
            pic_name_save = []
            for pic_dir in filename_bylist:
                pic_name_save.append(os.path.split(pic_dir)[1])
            gl.set_value('pic_name_save', pic_name_save)
            OCR(filename_bylist)
            write_to_excel_all()
            messagebox.showinfo("好啦！","推算完啦，辛苦等待！")
        else:
            pass



if __name__ == "__main__":
    gl._init()
    root = Tk() # 创建主窗口
    root.title("魏雨飞の标注软件") # 主窗口名称
    app = Application(master=root)
    root.resizable(False, False) # 规定窗口不可缩放
    root.mainloop() # 进入事件循环
