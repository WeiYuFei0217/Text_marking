from import_list import *
 
def results_demo():
    
    a = [{'save_path': 'Output\\ticket.jpg', 'data': [{'text': 'F067846', 'confidence': 0.9800475835800171, \
        'text_box_position': [[1106, 507], [1760, 518], [1760, 666], [1106, 655]]}, {'text': '检票：22', 'confidence': 0.9943496584892273, \
        'text_box_position': [[3184, 535], [3680, 535], [3680, 688], [3184, 688]]}, {'text': '北京南站', 'confidence': 0.999930202960968, \
        'text_box_position': [[1264, 693], [1994, 693], [1994, 879], [1264, 879]]}, {'text': '天津站', 'confidence': 0.9989504218101501, \
        'text_box_position': [[2837, 715], [3571, 715], [3571, 884], [2837, 884]]}, {'text': 'C2565', 'confidence': 0.9947913885116577, \
        'text_box_position': [[2157, 742], [2673, 731], [2673, 884], [2157, 895]]}, {'text': 'Beijingnan', 'confidence': 0.9975405931472778, \
        'text_box_position': [[1324, 884], [1875, 906], [1870, 1026], [1319, 1004]]}, {'text': 'Tianjin', 'confidence': 0.9982953667640686, \
        'text_box_position': [[2976, 890], [3363, 906], [3358, 1026], [2971, 1010]]}, {'text': '2019年04月03日09:36开', \
        'confidence': 0.9840608239173889, 'text_box_position': [[1160, 1021], [2509, 1037], [2509, 1157], [1160, 1141]]}, \
        {'text': '02车03C号', 'confidence': 0.9970527291297913, 'text_box_position': [[2777, 1043], [3338, 1043], [3338, 1163], [2777, 1163]]}, \
        {'text': '￥54.5元', 'confidence': 0.9837226271629333, 'text_box_position': [[1180, 1185], [1617, 1185], [1617, 1305], [1180, 1305]]}, \
        {'text': '网', 'confidence': 0.9999434947967529, 'text_box_position': [[2192, 1185], [2321, 1185], [2321, 1310], [2192, 1310]]}, \
        {'text': '二等座', 'confidence': 0.9980916976928711, 'text_box_position': [[2956, 1179], [3333, 1179], [3333, 1332], [2956, 1332]]}, \
        {'text': '限乘当日当次车', 'confidence': 0.9988431930541992, 'text_box_position': [[1150, 1316], [1959, 1332], [1959, 1469], [1150, 1452]]}, \
        {'text': '始发改签', 'confidence': 0.9973716735839844, 'text_box_position': [[1160, 1474], [1612, 1474], [1612, 1594], [1160, 1594]]}, \
        {'text': '2302051998****156X装喻丽', 'confidence': 0.9590480327606201, 'text_box_position': [[1140, 1616], [2777, 1632], [2777, 1780], [1140, 1764]]}, \
        {'text': '买票请到12306发货请到95306', 'confidence': 0.9994847774505615, 'text_box_position': [[1448, 1807], [2792, 1818], [2792, 1933], [1448, 1922]]}, \
        {'text': '中国铁路祝您旅途愉快', 'confidence': 0.9980210065841675, 'text_box_position': [[1612, 1944], [2604, 1955], [2604, 2058], [1612, 2048]]}, \
        {'text': '10010301110403F067846', 'confidence': 0.9951126575469971, 'text_box_position': [[1160, 2113], [2346, 2129], [2346, 2228], [1160, 2211]]}, \
        {'text': '北京南售', 'confidence': 0.9984771013259888, 'text_box_position': [[2296, 2129], [2772, 2129], [2772, 2233], [2296, 2233]]}]}]

    gl.set_value('results', a)
    gl.set_value('frozen', True)