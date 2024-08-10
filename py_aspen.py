import win32com.client as win32
import numpy as np
import time
import sys
import os
import psutil


def get_pid(process_name):
    for proc in psutil.process_iter():
        if proc.name() == process_name:
            return proc.pid


BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir, os.pardir, os.pardir))
sys.path.append(BASE_DIR)


class PyASPENPlus(object):
    """使用Python运行ASPEN模拟"""

    def __init__(self):
        self.app = None

    def init_app(self, ap_version: str = '10.0'):
        """开启ASPEN Plus

        :param ap_version: ASPEN Plus版本号, defaults to '10.0'
        """
        version_match = {
            '11.0': '37.0',
            '10.0': '36.0',
            '9.0': '35.0',
            '8.8': '34.0',
        }
        self.app = win32.Dispatch(f'Apwn.Document.{version_match[ap_version]}')

    def load_ap_file(self, file_name: str, file_dir: str = None, visible: bool = False, dialogs: bool = False):
        """载入待运行的ASPEN文件"""
        # 文件类型检查.
        if (not file_name.endswith('.apw')) and (not file_name.endswith('.bkp')):
            raise ValueError('not an valid ASPEN file')

        self.file_dir = os.getcwd() if file_dir is None else file_dir  # ASPEN文件所处目录, 默认为当前目录

        self.app.InitFromArchive2(os.path.join(self.file_dir, file_name))
        self.app.Visible = 1 if visible else 0
        self.app.SuppressDialogs = 0 if dialogs else 1

        print(f'The ASPEN file "{file_name}" has been reloaded')

    def assign_node_values(self, nodes: list, values: list, call_address: dict):
        for i, node in enumerate(nodes):
            if isinstance(values[i], (float, int)):  # 检查值是否是 float 或者 int 类型
                if -1e8 <= values[i] <= 1e8:  # 检查值是否在 -1e8 到 1e8 的范围内
                    try:
                        self.app.Tree.FindNode(call_address[node]).Value = values[i]
                        print(f"Successful input: {node}: {values[i]} ")
                    except Exception as e:
                        print(f"Error setting value for node '{node}' and Value '{values[i]}': {e}")
                else:
                    print(f"Value {values[i]} for node '{node}' is out of range (-1e8 to +1e8). Skipping assignment.")
            else:
                print(f"Value {values[i]} for node '{node}' is not a numeric type (float or int). Skipping assignment.")

    def assign_node_value1(self, value1: float, call_address1: str):
        self.app.Tree.FindNode(call_address1).Value = value1

    def run_simulation(self, reinit: bool = True, sleep: float = 2.0):
        """进行模拟

        :param reinit: 是否重新初始化迭代参数设置, defaults to True
        :param sleep: 每次检测运行状态的间隔时长, defaults to 2.0
        """
        if reinit:
            self.app.Reinit()

        self.app.Engine.Run2()
        while self.app.Engine.IsRunning == 1:
            time.sleep(sleep)

    def get_target_values(self, target_nodes: list, call_address) -> list:
        """从模拟结果中获得目标值"""
        values = []
        for node in target_nodes:
            values.append(self.app.Tree.FindNode(call_address[node]).Value)
        return values

    def get_target_value1(self, call_address1: str):
        """从模拟结果中获得目标值"""
        return self.app.Tree.FindNode(call_address1).Value

    def check_simulation_status(self) -> list:
        """检查模拟是否收敛等"""
        value = self.app.Tree.FindNode(r'\Data\Results Summary\Run-Status\Output\RUNID').Value
        file_path = os.path.join(self.file_dir, f'{value}.his')

        with open(file_path, 'r') as f:
            isError = np.any(np.array([line.find('SEVERE ERROR') for line in f.readlines()]) >= 0)
        return [not isError]

    def quit_app(self):
        self.app.Quit()

    def close_app(self):
        self.app.Close()

    def result_error(self):
        errall = ''
        errormessage = []
        node = self.app.Tree.FindNode(r"\Data\Results Summary\Run-Status\Output\PER_ERROR")
        if node is None:
            return 'error'
        else:
            for e in node.Elements:
                # print(e.Value)
                errormessage += e.value
                if '=' in errormessage:
                    break
            errall = errall.join(errormessage)

            if 'error' in errall:
                errall = 'error'
            else:
                errall = 'OK'
            return errall


# sample use
if __name__ == '__main__':

    # ---- 接口和值 ---------------------------------------------------------------------------------

    x_cols = ['FEED_pressure', 'FEED_ETHANOL', 'FEED_ACETIC', 'FEED_H2O']
    y_cols = ['PRODUCT_ETHYL-01']

    # 自行整理调用地址.
    # 调用地址查找方法参考：https://zhuanlan.zhihu.com/p/321125404
    call_address = {
        'FEED_pressure': r'\Data\Streams\FEED\Input\PRES\MIXED',
        'FEED_ETHANOL': r'\Data\Streams\FEED\Input\FLOW\MIXED\ETHANOL',
        'FEED_ACETIC': r'\Data\Streams\FEED\Input\FLOW\MIXED\ACETIC',
        'FEED_H2O': r'\Data\Streams\FEED\Input\FLOW\MIXED\H2O',

        'PRODUCT_ETHYL-01': r'\Data\Streams\PRODUCT\Output\MOLEFLOW\MIXED\ETHYL-01',
    }

    x_range = {
        'FEED_pressure': [0.1000, 0.1020],
        'FEED_ETHANOL': [200.0, 230.0],
        'FEED_ACETIC': [210.0, 240.0],
        'FEED_H2O': [710.0, 750.0],

        'PRODUCT_ETHYL-01': None,
    }


    # ---- ASPEN 模拟 ------------------------------------------------------------------------------

    def random_x_values():
        x_values = []
        for i, x_col in enumerate(x_cols):
            x_values.append(np.random.uniform(*x_range[x_col]))
        return x_values


    # 指定ASPEN文件名和所处目录.
    file_name = 'cstr.bkp'
    file_dir = os.getcwd()

    # 进行ASPEN模拟.
    pyaspen = PyASPENPlus()

    pyaspen.init_app()
    pyaspen.load_ap_file(file_name, file_dir)

    x_records, y_records, status_records = [], [], []
    repeats = 10
    for i in range(repeats):
        print(f'simulating {i}')

        # 随机给定一个参数值.
        x_values = random_x_values()

        pyaspen.assign_node_values(x_cols, x_values, call_address)
        pyaspen.run_simulation(reinit=False)

        y_values = pyaspen.get_target_values(y_cols, call_address)
        simul_status = pyaspen.check_simulation_status()

        x_records.append(x_values)
        y_records.append(y_values)
        status_records.append(simul_status)

    """
    pyaspen.close_app()

    process_name = "AspenPlus.exe"
    p = psutil.Process(get_pid(process_name))
    p.terminate()
    """
